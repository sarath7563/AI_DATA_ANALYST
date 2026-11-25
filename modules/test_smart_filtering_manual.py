import pandas as pd

from modules.smart_filtering import answer_filtered_question


df = pd.DataFrame({
    "Product": ["Laptop", "Phone", "Laptop", "Phone"],
    "Category": ["Electronics", "Electronics", "Electronics", "Electronics"],
    "Sales": [50000, 20000, 60000, 30000],
    "Profit": [10000, 4000, 12000, 6000],
})

questions = [
    "What is the total Sales?",
    "What is the average Profit?",
    "What is the highest Sales?",
    "What is the lowest Profit?",
    "Show Laptop records",
]

for question in questions:
    result = answer_filtered_question(df, question)

    print(f"\nQuestion: {question}")
    print(f"Success: {result['success']}")
    print(f"Answer: {result['answer']}")