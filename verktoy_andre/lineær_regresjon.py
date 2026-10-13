import pandas as pd

# Load the uploaded CSV file
file_path = "C:/Users/ifi/Downloads/(3) Utgående spor 2D.csv"

df = pd.read_csv(file_path)

# Display the first few rows to understand the structure
df.head()





# Rens kolonnenavn ved å fjerne ="" rundt navnene
df.columns = df.columns.str.replace(r'^="|"$', '', regex=True)

# Konverter tid til datetime
df["End Date (Local)(UTC+2)"] = pd.to_datetime(df["End Date (Local)(UTC+2)"])

# Filtrer ut kun "Value (m)" kolonner (hver tredje kolonne)
value_columns = [col for col in df.columns if "Value (m)" in col and "Min" not in col and "Max" not in col]

# Smelt sammen alle sensorverdier til én kolonne
df_melted = df.melt(id_vars=["End Date (Local)(UTC+2)"], value_vars=value_columns,
                    var_name="Sensor", value_name="Value (m)")

# Fjern rader med manglende verdier
df_melted.dropna(inplace=True)

# Konverter tid til sekunder siden start
df_melted["Time (s)"] = (df_melted["End Date (Local)(UTC+2)"] - df_melted["End Date (Local)(UTC+2)"].min()).dt.total_seconds()

# Lineær regresjon
X = df_melted[["Time (s)"]]
y = df_melted["Value (m)"]
model = LinearRegression()
model.fit(X, y)
df_melted["Predicted"] = model.predict(X)

# Plot
plt.figure(figsize=(12, 6))
sns.scatterplot(data=df_melted, x="End Date (Local)(UTC+2)", y="Value (m)", hue="Sensor", alpha=0.6, legend=False)
sns.lineplot(data=df_melted, x="End Date (Local)(UTC+2)", y="Predicted", color="black", label="Lineær regresjon")
plt.title("Samlet lineær regresjon av alle sensorer over tid")
plt.xlabel("Tid")
plt.ylabel("Verdi (m)")
plt.xticks(rotation=45)
plt.tight_layout()
plt.grid(True)
plt.show()