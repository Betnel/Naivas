import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
file_path = "Consumer Insights Specialist Case Study..xlsx"
xls = pd.ExcelFile(file_path)

# Load sales data (Task 2 Data)
df_sales = pd.read_excel(xls, sheet_name="Task 2 Data", engine="openpyxl")

# Load customer satisfaction data (Task 1 Data)
df_satisfaction = pd.read_excel(xls, sheet_name="Task 1 Data", engine="openpyxl", header=1)

# Convert TRANSDATE to datetime format
df_sales["TRANSDATE"] = pd.to_datetime(df_sales["TRANSDATE"], errors='coerce')

def basic_insights():
    # 1️⃣ Top 5 Best-Selling Items (by quantity sold)
    top_items = df_sales.groupby("ITEMNAME")["QTY"].sum().nlargest(5)
    
    # 2️⃣ Top 3 Brands Contributing Highest Revenue
    top_brands = df_sales.groupby("BRAND")["NETAMOUNTINCLTAX"].sum().nlargest(3)
    
    # 3️⃣ Day of the Week with Highest Sales Volume
    df_sales["DAY_OF_WEEK"] = df_sales["TRANSDATE"].dt.day_name()
    top_sales_day = df_sales.groupby("DAY_OF_WEEK")["QTY"].sum().idxmax()
    
    print("\nTop 5 Best-Selling Items:\n", top_items)
    print("\nTop 3 Brands by Revenue:\n", top_brands)
    print("\nDay with Highest Sales Volume:", top_sales_day)
    
    # Visualization with data labels
    plt.figure(figsize=(8, 5))
    ax1 = sns.barplot(x=top_items.values, y=top_items.index, palette="viridis")
    plt.title("Top 5 Best-Selling Items by Quantity Sold")
    plt.xlabel("Quantity Sold")
    plt.ylabel("Item Name")
    for i, v in enumerate(top_items.values):
        ax1.text(v + 0.5, i, f"{int(v)}", va='center')
    plt.savefig("top_selling_items.png")
    plt.show()
    
    plt.figure(figsize=(6, 4))
    ax2 = sns.barplot(x=top_brands.values, y=top_brands.index, palette="magma")
    plt.title("Top 3 Brands by Revenue")
    plt.xlabel("Total Revenue (KES)")
    plt.ylabel("Brand")
    for i, v in enumerate(top_brands.values):
        ax2.text(v + 0.5, i, f"{v:,.0f}", va='center')
    plt.savefig("top_brands.png")
    plt.show()
    
    sales_by_day = df_sales.groupby("DAY_OF_WEEK")["QTY"].sum().sort_values()
    plt.figure(figsize=(8, 5))
    ax3 = sns.barplot(x=sales_by_day.index, y=sales_by_day.values, palette="coolwarm")
    plt.title("Total Sales Volume by Day of the Week")
    plt.xlabel("Day of the Week")
    plt.ylabel("Total Quantity Sold")
    plt.xticks(rotation=45)
    for i, v in enumerate(sales_by_day.values):
        ax3.text(i, v + 0.5, f"{int(v)}", ha='center')
    plt.savefig("sales_by_day.png")
    plt.show()

def customer_satisfaction_analysis():
    df_satisfaction.rename(columns={
        "On a scale of 1-10, how likely are you to recommend STORE A  to your friends and family? (1 = Not likely at all, 10 = very likely)": "Recommendation_Score"
    }, inplace=True)
    
    df_satisfaction["Recommendation_Score"] = pd.to_numeric(df_satisfaction["Recommendation_Score"], errors='coerce')
    avg_recommendation_score = df_satisfaction["Recommendation_Score"].mean()
    satisfaction_counts = df_satisfaction["Recommendation_Score"].value_counts().sort_index()
    top_challenges = df_satisfaction["What challenges have you faced in STORE A?"].value_counts().head(5)
    
    # Visualizing Satisfaction Scores with data labels
    plt.figure(figsize=(8, 5))
    ax4 = sns.barplot(x=satisfaction_counts.index, y=satisfaction_counts.values)
    plt.title("Customer Satisfaction Score Distribution")
    plt.xlabel("Satisfaction Score (1-10)")
    plt.ylabel("Number of Customers")
    for i, v in enumerate(satisfaction_counts.values):
        ax4.text(i, v + 0.5, str(int(v)), ha='center')
    plt.savefig("customer_satisfaction.png")
    plt.show()
    
    print("\nAverage Recommendation Score:", avg_recommendation_score)
    print("\nTop Challenges Faced by Customers:\n", top_challenges)

# Run all analyses
basic_insights()
customer_satisfaction_analysis()
