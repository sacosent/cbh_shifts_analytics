import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

DATA_FILE = "Problems we tackle, Shift Offers v3.xlsx"

OUTPUT_DIR = "report_output"

# Column names as per glossary
COLS = [
    "SHIFT_ID",
    "WORKER_ID",
    "WORKPLACE_ID",
    "SHIFT_START_AT",
    "SHIFT_CREATED_AT",
    "OFFER_VIEWED_AT",
    "DURATION",
    "SLOT",
    "CLAIMED_AT",
    "DELETED_AT",
    "IS_VERIFIED",
    "CANCELED_AT",
    "IS_NCNS",
    "PAY_RATE",
    "CHARGE_RATE",
]

def load_data(path=DATA_FILE):
    """Load Excel dataset with proper dtypes."""
    df = pd.read_excel(path, engine="openpyxl")
    df.columns = COLS
    # Convert date columns
    date_cols = [
        "SHIFT_START_AT",
        "SHIFT_CREATED_AT",
        "OFFER_VIEWED_AT",
        "CLAIMED_AT",
        "DELETED_AT",
        "CANCELED_AT",
    ]
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def compute_summary(df):
    """Compute summary metrics for the dataset."""
    summary = {}
    summary["total_offers"] = len(df)
    summary["unique_workers"] = df["WORKER_ID"].nunique()
    summary["unique_shifts"] = df["SHIFT_ID"].nunique()
    summary["unique_workplaces"] = df["WORKPLACE_ID"].nunique()
    summary["claimed_offers"] = df["CLAIMED_AT"].notna().sum()
    summary["canceled_offers"] = df["CANCELED_AT"].notna().sum()
    summary["deleted_offers"] = df["DELETED_AT"].notna().sum()
    summary["no_show_offers"] = df["IS_NCNS"].sum()
    summary["worked_offers"] = df["IS_VERIFIED"].sum()
    return summary


def plot_pay_rate_distribution(df):
    sns.histplot(df["PAY_RATE"].dropna(), kde=True)
    plt.xlabel("Pay Rate")
    plt.title("Distribution of Pay Rate")
    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/pay_rate_distribution.png")
    plt.clf()


def plot_slot_counts(df):
    sns.countplot(x="SLOT", data=df)
    plt.xlabel("Slot")
    plt.ylabel("Count")
    plt.title("Shift Counts by Slot")
    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/slot_counts.png")
    plt.clf()


def plot_monthly_offers(df):
    df_month = df["SHIFT_START_AT"].dt.to_period("M").value_counts().sort_index()
    df_month.plot(kind="bar")
    plt.xlabel("Month")
    plt.ylabel("Number of Offers")
    plt.title("Offers per Month")
    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/offers_per_month.png")
    plt.clf()


def generate_report(summary):
    with open(f"{OUTPUT_DIR}/summary.md", "w") as f:
        f.write("# Shift Offer Summary\n\n")
        for key, value in summary.items():
            f.write(f"- **{key.replace('_', ' ').title()}**: {value}\n")
        f.write("\n")
        f.write("## Conclusions\n")
        f.write("These metrics provide insight into worker engagement, shift availability, "
                "and overall marketplace activity. High cancellation or no-show rates might "
                "indicate issues with shift fulfillment, while pay rate distribution helps "
                "identify competitive compensation ranges.\n")


def main():
    import os
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    df = load_data()
    summary = compute_summary(df)
    plot_pay_rate_distribution(df)
    plot_slot_counts(df)
    plot_monthly_offers(df)
    generate_report(summary)


if __name__ == "__main__":
    main()
