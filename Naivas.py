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
    
    # Visualization
    plt.figure(figsize=(8, 5))
    sns.barplot(x=top_items.values, y=top_items.index, palette="viridis")
    plt.title("Top 5 Best-Selling Items by Quantity Sold")
    plt.xlabel("Quantity Sold")
    plt.ylabel("Item Name")
    plt.show()
    
    plt.figure(figsize=(6, 4))
    sns.barplot(x=top_brands.values, y=top_brands.index, palette="magma")
    plt.title("Top 3 Brands by Revenue")
    plt.xlabel("Total Revenue (KES)")
    plt.ylabel("Brand")
    plt.show()
    
    sales_by_day = df_sales.groupby("DAY_OF_WEEK")["QTY"].sum().sort_values()
    plt.figure(figsize=(8, 5))
    sns.barplot(x=sales_by_day.index, y=sales_by_day.values, palette="coolwarm")
    plt.title("Total Sales Volume by Day of the Week")
    plt.xlabel("Day of the Week")
    plt.ylabel("Total Quantity Sold")
    plt.xticks(rotation=45)
    plt.show()

# Define the customer_analysis function
def customer_analysis():
    print("Customer analysis function is not yet implemented.")

# Run all analyses
basic_insights()
customer_analysis()
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
    
    # Visualization
    plt.figure(figsize=(8, 5))
    sns.barplot(x=top_items.values, y=top_items.index, palette="viridis")
    plt.title("Top 5 Best-Selling Items by Quantity Sold")
    plt.xlabel("Quantity Sold")
    plt.ylabel("Item Name")
    plt.show()
    
    plt.figure(figsize=(6, 4))
    sns.barplot(x=top_brands.values, y=top_brands.index, palette="magma")
    plt.title("Top 3 Brands by Revenue")
    plt.xlabel("Total Revenue (KES)")
    plt.ylabel("Brand")
    plt.show()
    
    sales_by_day = df_sales.groupby("DAY_OF_WEEK")["QTY"].sum().sort_values()
    plt.figure(figsize=(8, 5))
    sns.barplot(x=sales_by_day.index, y=sales_by_day.values, palette="coolwarm")
    plt.title("Total Sales Volume by Day of the Week")
    plt.xlabel("Day of the Week")
    plt.ylabel("Total Quantity Sold")
    plt.xticks(rotation=45)
    plt.show()

# Define the customer_satisfaction_analysis function
def customer_satisfaction_analysis():
    print("Customer satisfaction analysis function is not yet implemented.")

# Define the advanced_analysis function
def advanced_analysis():
    print("Advanced analysis function is not yet implemented.")

# Define the recommendations function
def recommendations():
    print("Recommendations function is not yet implemented.")

# Run all analyses
basic_insights()
customer_analysis()
advanced_analysis()
customer_satisfaction_analysis()
recommendations()
    