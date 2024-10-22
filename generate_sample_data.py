import pandas as pd
import random

# Generate sample data
data = {
    'name': [
        'John Doe', 'Jane Smith', 'Mike Johnson', 'Emily Brown', 'Chris Lee',
        'Sarah Wilson', 'David Taylor', 'Lisa Anderson', 'Tom Harris', 'Emma Davis'
    ],
    'email': [
        'john.doe@example.com', 'jane.smith@example.com', 'mike.j@example.com',
        'emily.b@example.com', 'chris.lee@example.com', 'sarah.w@example.com',
        'david.t@example.com', 'lisa.a@example.com', 'tom.h@example.com', 'emma.d@example.com'
    ],
    'batch': ['2020', '2021', '2019', '2022', '2020', '2021', '2019', '2022', '2020', '2021'],
    'ticket_number': [f'IT{random.randint(1000, 9999)}' for _ in range(10)],
    'payment_verified': ['Yes', 'Yes', 'No', 'Yes', 'Yes', 'No', 'Yes', 'Yes', 'Yes', 'No']
}

# Create DataFrame
df = pd.DataFrame(data)

# Save to Excel
df.to_excel('responses.xlsx', index=False)

print("Sample Excel file 'responses.xlsx' has been created.")