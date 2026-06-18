"""
Step 5: AI-generated reconciliation summary (with verification)
-----------------------------------------------------------------
Builds a prompt for ChatGPT/Claude based on the reconciliation results,
and provides a lightweight check that the numbers in the AI's response
match the actual data before anyone signs off on it.

Principle: the AI-generated draft is never submitted as-is. Every number
it produces is checked against the source data, and only a human-reviewed
version becomes the final report.
"""

import pandas as pd
from openpyxl import load_workbook


def build_prompt(summary_df: pd.DataFrame, result_df: pd.DataFrame) -> str:
    summary_text = summary_df.to_string(index=False)

    sample_rows = []
    for status in result_df["Status"].unique():
        sample = result_df[result_df["Status"] == status].head(2)
        sample_rows.append(sample.to_string(index=False))
    samples_text = "\n\n".join(sample_rows)

    prompt = f"""You are an accounting assistant.
Based on the reconciliation results below, write a short bank reconciliation summary.

Rules:
- Every number you mention MUST exactly match the numbers provided below. Do not estimate or round.
- Separate matched transactions, bank-only transactions, book-only transactions,
  possible duplicates, amount mismatches, and outstanding items into their own short paragraphs.
- Use a professional but simple tone suitable for a small business owner.
- End with a short "Recommended Next Steps" section (3 bullet points max), and explicitly
  call out that Amount Mismatch and Possible Duplicate items need human review before the
  period can be closed.

=== Status Summary ===
{summary_text}

=== Sample transactions per status ===
{samples_text}
"""
    return prompt


def verify_ai_summary(ai_text: str, summary_df: pd.DataFrame) -> list:
    """
    Checks whether every Count value from summary_df literally appears in
    the AI-generated text. This is not a complete verification -- it's a
    fast first pass to catch obvious numeric drift before a human reads
    the full draft line by line.
    """
    issues = []
    for _, row in summary_df.iterrows():
        status, count = row["Status"], int(row["Count"])
        if str(count) not in ai_text:
            issues.append(
                f"[NEEDS REVIEW] Count for '{status}' ({count}) was not found verbatim "
                f"in the AI summary text. Check manually."
            )
    if not issues:
        issues.append("Basic numeric check passed: all Summary Count values were found in the text.")
    return issues


if __name__ == "__main__":
    wb_path = "/home/claude/AI-Bank-Reconciliation/output/reconciliation_output.xlsx"
    result_df = pd.read_excel(wb_path, sheet_name="Reconciliation Detail")
    summary_df = pd.read_excel(wb_path, sheet_name="Summary").dropna(subset=["Status"])
    summary_df = summary_df[summary_df["Status"] != "Total Transactions"]

    prompt = build_prompt(summary_df, result_df)

    print("=" * 60)
    print("Copy the prompt below into ChatGPT / Claude:")
    print("=" * 60)
    print(prompt)

    with open("/home/claude/AI-Bank-Reconciliation/output/ai_prompt.txt", "w") as f:
        f.write(prompt)
    print("\nPrompt saved to output/ai_prompt.txt")
    print("\nAfter you get the AI's response, run verify_ai_summary(ai_text, summary_df)")
    print("as a first-pass numeric check, then read it yourself before using it as the")
    print("final report. See README.md > 'AI Usage & Verification' for the full process.")
