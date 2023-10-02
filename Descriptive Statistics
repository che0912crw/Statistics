import pandas as pd
import numpy as np
from scipy import stats
import plotly.graph_objects as go

# Load the Excel file into a DataFrame
file_path = 'Your File.xlsx'
sheet_name = 'Sheet_name'
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Exclude the first column from calculations
data = df.iloc[:, 1:]

# Initialize lists to store the statistics
statistics = ["Valid Values", "Missing Values", "Mean", "Median", "Variance", "Standard Deviation", "Maximum", "Minimum", "Range",
              "Q1 (25th Percentile)", "Q3 (75th Percentile)", "Interquartile Range", "Skewness",
              "Standard Error of Skewness", "Kurtosis", "Standard Error of Kurtosis", "Shapiro-Wilk p-value",
              "Shapiro-Wilk Test Result"]

# Initialize empty lists to store results for each statistic
results = [[] for _ in statistics]

# Loop through each column and calculate the statistics
for column in df.columns[1:]:  # Skip the first column
    valid_values = data[column].count()
    missing_values = len(data[column]) - valid_values
    
    mean = np.mean(data[column])
    median = np.median(data[column])
    variance = np.var(data[column])
    std_dev = np.std(data[column])
    max_value = np.max(data[column])
    min_value = np.min(data[column])
    range_value = max_value - min_value
    q1 = np.percentile(data[column], 25)
    q3 = np.percentile(data[column], 75)
    iqr = q3 - q1
    skewness = stats.skew(data[column])
    se_skewness = stats.sem(data[column], axis=None, ddof=1)
    kurtosis = stats.kurtosis(data[column])
    n = len(data[column])
    se_kurtosis = np.sqrt((6 * (n - 2)) / ((n + 1) * (n + 3)))
    shapiro_stat, shapiro_p = stats.shapiro(data[column])
    shapiro_result = "Not Normally Distributed" if shapiro_p < 0.05 else "Normally Distributed"

    # Append the rounded results to the corresponding lists
    results[0].append(valid_values)
    results[1].append(missing_values)
    results[2].append(round(mean, 3))
    results[3].append(round(median, 3))
    results[4].append(round(variance, 3))
    results[5].append(round(std_dev, 3))
    results[6].append(round(max_value, 3))
    results[7].append(round(min_value, 3))
    results[8].append(round(range_value, 3))
    results[9].append(round(q1, 3))
    results[10].append(round(q3, 3))
    results[11].append(round(iqr, 3))
    results[12].append(round(skewness, 3))
    results[13].append(round(se_skewness, 3))
    results[14].append(round(kurtosis, 3))
    results[15].append(round(se_kurtosis, 3))
    results[16].append(round(shapiro_p, 3))
    results[17].append(shapiro_result)

# Create a DataFrame for the results
result_df = pd.DataFrame(data=results, columns=df.columns[1:])

# Define colors for alternating row background
colors = ['lightgray' if i % 2 == 0 else 'white' for i in range(len(result_df))]

# Create a Plotly table
fig = go.Figure(data=[go.Table(
    header=dict(values=["Statistic"] + result_df.columns.tolist(),
                fill_color='lightblue',
                align='center'),
    cells=dict(values=[statistics] + result_df.values.T.tolist(),
               fill_color=[colors] + ['white']*len(result_df),
               align='center'))
])

# Show the Plotly table
fig.show()


##Credits for Chat GPT
