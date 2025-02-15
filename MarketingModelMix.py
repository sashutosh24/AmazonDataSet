# Import required libraries
import pandas as pd
import statsmodels.api as sm
from google.colab import files

# Step 1: Upload your dataset
uploaded = files.upload()  # Upload the dataset
file_name = list(uploaded.keys())[0]
data = pd.ExcelFile(file_name)

# Step 2: Load the relevant sheet
df = data.parse('Sheet1')  # Replace 'Sheet1' with the sheet name in your file

# Step 3: Inspect the dataset
print("Dataset Preview:")
print(df.head())

# Step 4: Prepare data for the Marketing Mix Model
# Ensure your dataset has columns for sales and marketing channel spends (e.g., TV, digital, radio)
# Example columns: ['Sales', 'TV_Spend', 'Digital_Spend', 'Radio_Spend', 'Promotions']

# Check for missing data
print("Missing values per column:\n", df.isnull().sum())

# Drop rows with missing values
df = df.dropna()

# Define dependent and independent variables
X = df[['TV_Spend', 'Digital_Spend', 'Radio_Spend', 'Promotions']]  # Replace with your actual column names
y = df['Sales']  # Replace with your sales column name

# Add a constant to the independent variables (for the intercept)
X = sm.add_constant(X)

# Step 5: Fit the Marketing Mix Model (Multiple Linear Regression)
model = sm.OLS(y, X).fit()

# Step 6: Display the model summary
print(model.summary())

# Step 7: Contribution of Marketing Channels
# Calculate channel contributions
coefficients = model.params[1:]  # Exclude the intercept
channel_spend = X.iloc[:, 1:].mean()  # Average spend for each channel
contribution = coefficients * channel_spend
contribution_percentage = (contribution / contribution.sum()) * 100

# Display contributions
print("\nChannel Contributions to Sales:")
for channel, contrib in zip(X.columns[1:], contribution_percentage):
    print(f"{channel}: {contrib:.2f}%")

# Step 8: Save the results
contribution_df = pd.DataFrame({
    'Channel': X.columns[1:],  # Marketing channels
    'Contribution (%)': contribution_percentage
})
contribution_df.to_excel("Marketing_Mix_Contributions.xlsx", index=False)
print("Channel contributions saved to Marketing_Mix_Contributions.xlsx")
