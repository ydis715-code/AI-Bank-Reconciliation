"""
AI-Assisted Bank Reconciliation Automation
--------------------------------------------
Compares bank transactions (bank_transactions.csv) against the company's
general ledger (book_transactions.csv) and classifies each transaction into
one of the standard bank reconciliation categories below.

    - Matched              : Amount, date, and payee reasonably agree
    - Amount Mismatch       : Same payee/date, but the amount differs
                              (e.g. data entry error)
    - Outstanding Item      : Recorded in the books but the date the bank
                              cleared it is 1-3 days different (standard
                              timing difference, e.g. outstanding check or
                              deposit in transit)
    - Possible Duplicate    : The same amount + date appears more than once
                              on the bank side (duplicate payment risk)
    - Bank Only             : Exists on the bank statement but has no
                              matching entry in the books (e.g. bank fees,
                              unrecorded charges)
    - Book Only             : Exists in the books but has no matching entry
                              on the bank statement (e.g. outstanding check
                              not yet cleared, deposit in transit)

Design notes
------------
1. A naive merge on Amount alone breaks when multiple transactions share
   the same amount (cartesian product problem). Instead, each bank
   transaction is scored against unmatched book transactions using a
   weighted combination of payee name similarity (rapidfuzz), date
   proximity, and amount similarity, and the single best-scoring
   candidate is selected as a 1:1 match.
2. AUTHOR'S NOTE (kept intentionally): the first version of
   find_best_match() had a bug where it returned the values from the
   LAST candidate evaluated in the loop instead of the candidate with the
   HIGHEST score. This produced ~20 false "Amount Mismatch" rows on a
   25-row test set. The fix was to store name_score/date_diff/amount_diff
   alongside best_score every time best_score is updated, not just at the
   end of the loop. This is exactly the kind of error a human reviewer
   needs to catch before trusting AI-generated logic -- see README.md,
   "AI Usage & Verification" section, for the full verification log.
3. Possible Duplicate is checked independently of match status, by
   grouping bank transactions on (Amount, Date) and flagging any group
   with more than one row.
"""

import pandas as pd
from rapidfuzz import fuzz


# ----------------------------
# 0. Configuration (adjust as needed for your data)
# ----------------------------
NAME_SIMILARITY_THRESHOLD = 60   # Minimum payee-name similarity score (0-100)
TIMING_DIFF_MAX_DAYS = 3         # Max date difference still treated as a timing difference
AMOUNT_MISMATCH_MAX_DIFF = 5.00  # Amount differences within this range are flagged
                                  # as likely data-entry errors rather than unrelated transactions


def normalize_name(name: str) -> str:
    """Normalize a payee/description string for comparison."""
    return str(name).upper().strip()


def load_data(bank_path: str, book_path: str):
    bank = pd.read_csv(bank_path, parse_dates=["Date"])
    book = pd.read_csv(book_path, parse_dates=["Date"])

    bank["Description_norm"] = bank["Description"].apply(normalize_name)
    book["Vendor_norm"] = book["Vendor/Customer"].apply(normalize_name)

    return bank, book


def find_best_match(bank_row, book_df, used_book_refs):
    """
    For a single bank transaction, find the most plausible candidate among
    book transactions not yet claimed by another match. Returns the best
    candidate even if it isn't a perfect match -- the caller decides
    whether the score is high enough to count as a real match.
    """
    candidates = book_df[~book_df["Book Ref"].isin(used_book_refs)]
    if candidates.empty:
        return None

    best_score = -1
    best_idx = None
    best_name_score = None
    best_date_diff = None
    best_amount_diff = None

    for idx, book_row in candidates.iterrows():
        name_score = fuzz.token_sort_ratio(
            bank_row["Description_norm"], book_row["Vendor_norm"]
        )

        date_diff = abs((bank_row["Date"] - book_row["Date"]).days)
        # Closer dates score higher (0 days = 100, 7+ days = 0)
        date_score = max(0, 100 - (date_diff * 15))

        amount_diff = abs(bank_row["Amount"] - book_row["Amount"])
        # Closer amounts score higher (exact match = 100)
        amount_score = 100 if amount_diff == 0 else max(0, 100 - amount_diff * 2)

        # Weighted total: name 40% + date 30% + amount 30%
        total_score = name_score * 0.4 + date_score * 0.3 + amount_score * 0.3

        # IMPORTANT: store the metrics for THIS candidate whenever it
        # becomes the new best, not the metrics from the final loop
        # iteration. (This is the line that fixed the bug described above.)
        if total_score > best_score:
            best_score = total_score
            best_idx = idx
            best_name_score = name_score
            best_date_diff = date_diff
            best_amount_diff = amount_diff

    if best_idx is None:
        return None

    return best_idx, best_score, best_name_score, best_date_diff, best_amount_diff


def classify_transactions(bank: pd.DataFrame, book: pd.DataFrame) -> pd.DataFrame:
    results = []
    used_book_refs = set()

    # ----------------------------
    # 1. For each bank transaction, find the most plausible book match
    # ----------------------------
    for _, bank_row in bank.iterrows():
        match = find_best_match(bank_row, book, used_book_refs)

        if match is None:
            results.append({
                "Bank Ref": bank_row["Bank Ref"],
                "Book Ref": None,
                "Date (Bank)": bank_row["Date"].date(),
                "Date (Book)": None,
                "Description / Vendor": bank_row["Description"],
                "Amount (Bank)": bank_row["Amount"],
                "Amount (Book)": None,
                "Status": "Bank Only",
                "Note": "No corresponding entry found in the books",
            })
            continue

        best_idx, total_score, name_score, date_diff, amount_diff = match
        book_row = book.loc[best_idx]

        # If even the best candidate scores too low, treat as no match
        if total_score < 50:
            results.append({
                "Bank Ref": bank_row["Bank Ref"],
                "Book Ref": None,
                "Date (Bank)": bank_row["Date"].date(),
                "Date (Book)": None,
                "Description / Vendor": bank_row["Description"],
                "Amount (Bank)": bank_row["Amount"],
                "Amount (Book)": None,
                "Status": "Bank Only",
                "Note": "No corresponding entry found in the books",
            })
            continue

        # Candidate confirmed -- remove from the pool so it can't be reused
        used_book_refs.add(book_row["Book Ref"])

        if amount_diff > 0 and amount_diff <= AMOUNT_MISMATCH_MAX_DIFF:
            status = "Amount Mismatch"
            note = f"Amount differs by {amount_diff:.2f} (possible data entry error)"
        elif amount_diff > AMOUNT_MISMATCH_MAX_DIFF:
            status = "Amount Mismatch"
            note = f"Amount differs by {amount_diff:.2f} (needs review)"
        elif date_diff > 0 and date_diff <= TIMING_DIFF_MAX_DAYS:
            status = "Outstanding Item"
            note = f"Date differs by {date_diff} day(s) (standard clearing delay)"
        elif date_diff > TIMING_DIFF_MAX_DAYS:
            status = "Outstanding Item"
            note = f"Date differs by {date_diff} day(s) (needs review)"
        else:
            status = "Matched"
            note = "Amount, date, and payee agree"

        results.append({
            "Bank Ref": bank_row["Bank Ref"],
            "Book Ref": book_row["Book Ref"],
            "Date (Bank)": bank_row["Date"].date(),
            "Date (Book)": book_row["Date"].date(),
            "Description / Vendor": bank_row["Description"],
            "Amount (Bank)": bank_row["Amount"],
            "Amount (Book)": book_row["Amount"],
            "Status": status,
            "Note": note,
        })

    # ----------------------------
    # 2. Remaining unmatched book transactions -> Book Only
    # ----------------------------
    unmatched_book = book[~book["Book Ref"].isin(used_book_refs)]
    for _, book_row in unmatched_book.iterrows():
        results.append({
            "Bank Ref": None,
            "Book Ref": book_row["Book Ref"],
            "Date (Bank)": None,
            "Date (Book)": book_row["Date"].date(),
            "Description / Vendor": book_row["Vendor/Customer"],
            "Amount (Bank)": None,
            "Amount (Book)": book_row["Amount"],
            "Status": "Book Only",
            "Note": "No corresponding entry found on the bank statement",
        })

    result_df = pd.DataFrame(results)

    # ----------------------------
    # 3. Flag Possible Duplicate independently of match status
    #    (same Amount + Date appearing more than once on the bank side)
    # ----------------------------
    bank_amount_date_counts = bank.groupby(["Amount", "Date"]).size()
    for i, row in result_df.iterrows():
        if row["Amount (Bank)"] is not None and row["Date (Bank)"] is not None:
            key = (row["Amount (Bank)"], pd.Timestamp(row["Date (Bank)"]))
            if key in bank_amount_date_counts.index and bank_amount_date_counts[key] > 1:
                result_df.at[i, "Status"] = "Possible Duplicate"
                result_df.at[i, "Note"] = "Same amount and date appears more than once on the bank statement"

    return result_df


def build_summary(result_df: pd.DataFrame) -> pd.DataFrame:
    summary = result_df["Status"].value_counts().reset_index()
    summary.columns = ["Status", "Count"]
    return summary


def build_reconciliation_statement(bank: pd.DataFrame, book: pd.DataFrame, result_df: pd.DataFrame) -> dict:
    """
    Builds the figures for a standard Bank Reconciliation Statement:

        Book Balance (ending)
          + Deposits in Transit (Book Only, positive amounts)
          - Outstanding Checks  (Book Only, negative amounts)
          = Adjusted Book Balance

        Bank Balance (ending)
          + Deposits in Transit not yet cleared by the bank
          - Outstanding Checks not yet cleared by the bank
          - Bank Service Charges / fees not yet recorded in the books
          = Adjusted Bank Balance

    Adjusted Book Balance should equal Adjusted Bank Balance.
    This function returns the components so they can be written into
    the Excel statement with live formulas rather than hardcoded totals.
    """
    book_ending_balance = book["Amount"].sum()
    bank_ending_balance = bank["Amount"].sum()

    book_only = result_df[result_df["Status"] == "Book Only"]
    deposits_in_transit = book_only[book_only["Amount (Book)"] > 0]["Amount (Book)"].sum()
    outstanding_checks = book_only[book_only["Amount (Book)"] < 0]["Amount (Book)"].sum()

    bank_only = result_df[result_df["Status"] == "Bank Only"]
    unrecorded_bank_items = bank_only["Amount (Bank)"].sum()

    return {
        "book_ending_balance": book_ending_balance,
        "bank_ending_balance": bank_ending_balance,
        "deposits_in_transit": deposits_in_transit,
        "outstanding_checks": outstanding_checks,
        "unrecorded_bank_items": unrecorded_bank_items,
    }


def main():
    bank, book = load_data(
        "/home/claude/AI-Bank-Reconciliation/data/bank_transactions.csv",
        "/home/claude/AI-Bank-Reconciliation/data/book_transactions.csv",
    )

    result_df = classify_transactions(bank, book)
    summary_df = build_summary(result_df)
    statement = build_reconciliation_statement(bank, book, result_df)

    print("=== Reconciliation Summary ===")
    print(summary_df.to_string(index=False))
    print("\n=== Reconciliation Statement Inputs ===")
    for k, v in statement.items():
        print(f"{k}: {v:.2f}")

    return result_df, summary_df, statement


if __name__ == "__main__":
    main()
