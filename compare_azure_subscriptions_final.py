import pandas as pd

# File paths
yesterday_file = 'azure_subscriptions_yesterday.xlsx'
today_file = 'azure_subscriptions_today.xlsx'

# Load Excel files
df_yesterday = pd.read_excel(yesterday_file)
df_today = pd.read_excel(today_file)

# Normalize Subscription IDs
yesterday_ids = set(df_yesterday['Subscription ID'].astype(str).str.strip())
today_ids = set(df_today['Subscription ID'].astype(str).str.strip())

# Identify differences
new_ids = today_ids - yesterday_ids
deleted_ids = yesterday_ids - today_ids

# Extract rows for new and deleted subscriptions
df_new = df_today[df_today['Subscription ID'].isin(new_ids)].copy()
df_deleted = df_yesterday[df_yesterday['Subscription ID'].isin(deleted_ids)].copy()

# Add status column
df_new.loc[:, "Status"] = "New"
df_deleted.loc[:, "Status"] = "Deleted"

# Combine results
df_result = pd.concat([df_new, df_deleted])

# Save to Excel
df_result.to_excel("subscription_changes.xlsx", index=False)

# Optional console summary
print(f"✅ Comparison completed.")
print(f"🆕 New subscriptions: {len(df_new)}")
print(f"❌ Deleted subscriptions: {len(df_deleted)}")
print(f"📄 Results saved to 'subscription_changes.xlsx'")
