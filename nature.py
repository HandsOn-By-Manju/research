import pandas as pd

# Create a list to store rows of data
data = []

# Loop to generate 50 rows with 4 columns
for i in range(1, 51):
    bird_name = f"Bird {i}"
    animal_name = f"Animal {i}"
    plant_name = f"Plant {i}"
    food_item = f"Food {i}"
    data.append([bird_name, animal_name, plant_name, food_item])

# Convert the list to a DataFrame with four columns
df = pd.DataFrame(data, columns=["Bird Name", "Animal Name", "Plant Name", "Food Item"])

# Export to Excel
df.to_excel("Nature_Data.xlsx", index=False)

print("Excel file 'Nature_Data.xlsx' created successfully!")
