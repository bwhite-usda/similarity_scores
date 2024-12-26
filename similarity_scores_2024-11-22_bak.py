import pandas as pd
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from itertools import product  # Used to generate combinations

# Load the Excel file
input_file = "ivntest.xlsx"
output_file = "similarity_scores.xlsx"

# Read the Excel file into a DataFrame
df = pd.read_excel(input_file)

# Check if required columns exist
if "Enabling Component Description" not in df.columns or "Dependent Component Description" not in df.columns:
    raise ValueError("The input file must contain columns 'Enabling Component Description' and 'Dependent Component Description'.")

# Fill missing values with an empty string to avoid issues with NaN
df = df.fillna("")

# Separate Enabling and Dependent components
enabling_df = df[[
    "Enabling Component Description",
    "Enabling Component",
    "Enabling Source",
    "Enabling Component URL",
    "Enabling Source Agency"
]].drop_duplicates()

dependent_df = df[[
    "Dependent Component Description",
    "Dependent Component",
    "Dependent Source",
    "Dependent Component URL",
    "Dependent Source Agency"
]].drop_duplicates()

# Generate all possible combinations of Enabling and Dependent rows
combinations = product(enabling_df.itertuples(index=False), dependent_df.itertuples(index=False))

# Create a new DataFrame for the combinations
combined_data = []
for enabling, dependent in combinations:
    combined_data.append({
        "Enabling Source": enabling[2],
        "Enabling Component": enabling[1],
        "Enabling Component Description": enabling[0],
        "Dependent Component": dependent[1],
        "Dependent Component Description": dependent[0],
        "Dependent Source": dependent[2],
        "Linkage mandated by what US Code or OMB policy?": "",  # You can add specific logic to fill this
        "Enabling Component URL": enabling[3],
        "Dependent Component URL": dependent[3],
        "Enabling Source Agency": enabling[4],
        "Dependent Source Agency": dependent[4],
        "Notes and keywords": "",  # You can add specific logic to fill this
        "Keywords Tab Items Found": "",  # You can add specific logic to fill this
        "Enabling Component Responsible Office": "",  # You can add specific logic to fill this
        "Dependent Component Responsible Office": "",  # You can add specific logic to fill this
    })

# Create a DataFrame for the combined data
combined_df = pd.DataFrame(combined_data)

# Ensure the columns match the sequence in the original `ivntest.xlsx`
original_columns = [
    "Enabling Source",
    "Enabling Component",
    "Enabling Component Description",
    "Dependent Component",
    "Dependent Component Description",
    "Dependent Source",
    "Linkage mandated by what US Code or OMB policy?",
    "Enabling Component URL",
    "Dependent Component URL",
    "Enabling Source Agency",
    "Dependent Source Agency",
    "Notes and keywords",
    "Keywords Tab Items Found",
    "Enabling Component Responsible Office",
    "Dependent Component Responsible Office"
]

# Reorder the columns to match the original sequence
combined_df = combined_df[original_columns]

# Function to calculate cosine similarity
def calculate_similarity(text1, text2):
    if not text1 or not text2:  # Avoid processing empty text
        return 0
    vectorizer = CountVectorizer().fit_transform([text1, text2])
    vectors = vectorizer.toarray()
    return cosine_similarity(vectors)[0, 1]

# Calculate similarity for each row
combined_df["Similarity"] = combined_df.apply(
    lambda row: calculate_similarity(row["Enabling Component Description"], row["Dependent Component Description"]), axis=1
)

# Save the updated DataFrame to a new Excel file
combined_df.to_excel(output_file, index=False, float_format="%.4f")

print(f"Similarity scores added and saved to {output_file}")
