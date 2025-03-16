import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.pyplot as plt
import os

def pie_graph(table, column_name, output_path):
    df = pd.DataFrame(table[1:], columns=table[0])
    data = df[column_name].value_counts()
    
    plt.figure(figsize=(8, 8))
    plt.pie(data, labels=data.index, autopct='%1.1f%%', startangle=140)
    plt.title(f'Pie Chart of {column_name}')
    
    # Save the pie chart as an image file
    plt.savefig(output_path)
    plt.close()