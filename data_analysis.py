import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
import numpy as np

class DataAnalyzer:
    def __init__(self, excel_path):
        self.df = pd.read_excel(excel_path, engine='openpyxl')
        self.output_dir = Path('visualizations')
        self.output_dir.mkdir(exist_ok=True)
        plt.style.use('default')  # Changed from 'seaborn' to 'default'
        
    def generate_all_charts(self):
        self._create_bar_chart()
        self._create_line_chart()
        self._create_pie_chart()
        self._create_area_chart()
        self._create_scatter_chart()
        self._create_column_chart()
        self._create_bubble_chart()
        self._create_histogram()
        self._create_box_plot()
        self._create_funnel_chart()
        self._create_waterfall_chart()
        
    def generate_statistics(self):
        stats = {
            'summary': self.df.describe(),
            'correlations': self.df.select_dtypes(include=[np.number]).corr(),
            'missing_values': self.df.isnull().sum()
        }
        return stats

    def _save_plot(self, name):
        plt.savefig(self.output_dir / f"{name}.png")
        plt.close()

    def _create_bar_chart(self):
        plt.figure(figsize=(10, 6))
        self.df.value_counts().head(10).plot(kind='bar')
        plt.title('Top 10 Categories Distribution')
        plt.tight_layout()
        self._save_plot('bar_chart')

    def _create_line_chart(self):
        plt.figure(figsize=(10, 6))
        if 'date' in self.df.columns:
            self.df.set_index('date').plot(kind='line')
            plt.title('Trend Over Time')
        plt.tight_layout()
        self._save_plot('line_chart')

    def _create_pie_chart(self):
        plt.figure(figsize=(10, 6))
        self.df.value_counts().head(5).plot(kind='pie', autopct='%1.1f%%')
        plt.title('Top 5 Categories Distribution')
        plt.tight_layout()
        self._save_plot('pie_chart')

    def _create_area_chart(self):
        plt.figure(figsize=(10, 6))
        if 'date' in self.df.columns:
            self.df.set_index('date').plot(kind='area', stacked=True)
            plt.title('Cumulative Trends')
        plt.tight_layout()
        self._save_plot('area_chart')

    def _create_scatter_chart(self):
        plt.figure(figsize=(10, 6))
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) >= 2:
            plt.scatter(self.df[numeric_cols[0]], self.df[numeric_cols[1]])
            plt.xlabel(numeric_cols[0])
            plt.ylabel(numeric_cols[1])
            plt.title('Correlation Scatter Plot')
        plt.tight_layout()
        self._save_plot('scatter_chart')

    def _create_column_chart(self):
        plt.figure(figsize=(10, 6))
        self.df.value_counts().head(10).plot(kind='bar')
        plt.title('Category Distribution')
        plt.tight_layout()
        self._save_plot('column_chart')

    def _create_bubble_chart(self):
        plt.figure(figsize=(10, 6))
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) >= 3:
            plt.scatter(self.df[numeric_cols[0]], 
                       self.df[numeric_cols[1]], 
                       s=self.df[numeric_cols[2]]*100)
            plt.xlabel(numeric_cols[0])
            plt.ylabel(numeric_cols[1])
            plt.title('Bubble Chart')
        plt.tight_layout()
        self._save_plot('bubble_chart')

    def _create_histogram(self):
        plt.figure(figsize=(10, 6))
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            plt.hist(self.df[numeric_cols[0]], bins=30)
            plt.title(f'Distribution of {numeric_cols[0]}')
        plt.tight_layout()
        self._save_plot('histogram')

    def _create_box_plot(self):
        plt.figure(figsize=(10, 6))
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            sns.boxplot(data=self.df[numeric_cols])
            plt.title('Statistical Distribution')
        plt.tight_layout()
        self._save_plot('box_plot')

    def _create_funnel_chart(self):
        plt.figure(figsize=(10, 6))
        values = self.df.value_counts().sort_values()
        plt.plot(values, range(len(values)), values)
        plt.title('Funnel Chart')
        plt.tight_layout()
        self._save_plot('funnel_chart')

    def _create_waterfall_chart(self):
        plt.figure(figsize=(10, 6))
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            values = self.df[numeric_cols[0]].head(10)
            cumsum = values.cumsum()
            plt.bar(range(len(values)), values)
            plt.plot(range(len(values)), cumsum, 'r-')
            plt.title('Waterfall Chart')
        plt.tight_layout()
        self._save_plot('waterfall_chart')

if __name__ == "__main__":
    analyzer = DataAnalyzer('your_data.xlsx')
    analyzer.generate_all_charts()
    stats = analyzer.generate_statistics()
