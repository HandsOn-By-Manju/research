import pandas as pd

# Load files
df_yesterday = pd.read_excel("azure_subscriptions_yesterday.xlsx")
df_today = pd.read_excel("azure_subscriptions_today.xlsx")

# Compare based on Subscription ID
yesterday_ids = set(df_yesterday['Subscription ID'].astype(str).str.strip())
today_ids = set(df_today['Subscription ID'].astype(str).str.strip())

new_ids = today_ids - yesterday_ids
deleted_ids = yesterday_ids - today_ids

# Filter rows for output
df_new = df_today[df_today['Subscription ID'].isin(new_ids)]
df_deleted = df_yesterday[df_yesterday['Subscription ID'].isin(deleted_ids)]

# Add status for clarity
df_new["Status"] = "New"
df_deleted["Status"] = "Deleted"

# Combine and save result
df_result = pd.concat([df_new, df_deleted])
df_result.to_excel("subscription_changes.xlsx", index=False)

print("Comparison completed. Output saved to 'subscription_changes.xlsx'.")
