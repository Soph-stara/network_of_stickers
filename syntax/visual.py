import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import os
from collections import Counter

def analyze_topic_frequencies(file_path):
    """Analyze and visualize topic frequencies from the data"""
    # Read the Excel file
    df = pd.read_excel(file_path)
    
    # Clean the topic data
    df['topic'] = df['topic'].astype(str).str.strip()
    
    # Remove N/A and empty values
    df = df[df['topic'].notna() & (df['topic'] != 'N/A') & (df['topic'] != 'nan')]
    
    # Count topic frequencies
    topic_counts = Counter(df['topic'])
    
    # Convert to DataFrame for easier plotting
    topic_df = pd.DataFrame({
        'Topic': list(topic_counts.keys()),
        'Count': list(topic_counts.values())
    })
    
    # Sort by count (descending)
    topic_df = topic_df.sort_values('Count', ascending=False).reset_index(drop=True)
    
    return topic_df

def create_bar_chart(topic_df, output_file):
    """Create a horizontal bar chart of topic frequencies"""
    plt.figure(figsize=(12, 10))
    
    # Set a nice color palette
    colors = sns.color_palette("viridis", len(topic_df))
    
    # Create horizontal bar chart
    bars = plt.barh(topic_df['Topic'], topic_df['Count'], color=colors)
    
    # Add count labels to the end of each bar
    for bar in bars:
        width = bar.get_width()
        plt.text(width + 0.3, bar.get_y() + bar.get_height()/2, 
                 f'{width:.0f}', ha='left', va='center', fontweight='bold')
    
    # Add labels and title
    plt.xlabel('Number of Mentions', fontsize=12)
    plt.ylabel('Topic', fontsize=12)
    plt.title('Frequency of Topics Mentioned', fontsize=16)
    
    # Remove top and right spines
    plt.gca().spines['top'].set_visible(False)
    plt.gca().spines['right'].set_visible(False)
    
    # Add gridlines
    plt.grid(axis='x', linestyle='--', alpha=0.7)
    
    # Tight layout
    plt.tight_layout()
    
    # Save the chart
    plt.savefig(output_file, dpi=300, bbox_inches='tight')
    plt.close()
    
    print(f"Bar chart saved as {output_file}")

def create_pie_chart(topic_df, output_file):
    """Create a pie chart of topic frequencies"""
    # For pie chart, let's group less frequent topics as "Other"
    threshold = 3  # Topics with less than this count will be grouped
    
    # Copy the dataframe to avoid modifying the original
    pie_df = topic_df.copy()
    
    # Identify topics to group
    to_group = pie_df[pie_df['Count'] < threshold]
    
    if not to_group.empty:
        # Create a new row for "Other" category
        other_count = to_group['Count'].sum()
        
        # Remove the rows to be grouped
        pie_df = pie_df[pie_df['Count'] >= threshold]
        
        # Add the "Other" row
        other_row = pd.DataFrame({'Topic': ['Other'], 'Count': [other_count]})
        pie_df = pd.concat([pie_df, other_row], ignore_index=True)
    
    # Create the pie chart
    plt.figure(figsize=(12, 10))
    
    # Use a nice colormap
    colors = sns.color_palette("viridis", len(pie_df))
    
    # Plot
    wedges, texts, autotexts = plt.pie(
        pie_df['Count'], 
        labels=pie_df['Topic'], 
        autopct='%1.1f%%',
        startangle=90,
        colors=colors,
        shadow=False
    )
    
    # Make percentage labels more readable
    for autotext in autotexts:
        autotext.set_fontsize(10)
        autotext.set_fontweight('bold')
        autotext.set_color('white')
    
    # Equal aspect ratio ensures the pie chart is circular
    plt.axis('equal')
    
    plt.title('Distribution of Topics', fontsize=16)
    
    # Tight layout
    plt.tight_layout()
    
    # Save the chart
    plt.savefig(output_file, dpi=300, bbox_inches='tight')
    plt.close()
    
    print(f"Pie chart saved as {output_file}")

def create_wordcloud(topic_df, output_file):
    """Create a word cloud visualization of topics"""
    try:
        from wordcloud import WordCloud
        
        # Create a dictionary with topic as key and frequency as value
        topic_dict = dict(zip(topic_df['Topic'], topic_df['Count']))
        
        # Generate the word cloud
        wordcloud = WordCloud(
            width=800, 
            height=400, 
            background_color='white',
            colormap='viridis',
            max_words=100,
            min_font_size=10,
            max_font_size=100,
            random_state=42
        ).generate_from_frequencies(topic_dict)
        
        # Plot
        plt.figure(figsize=(10, 8))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        plt.tight_layout(pad=0)
        
        # Save the word cloud
        plt.savefig(output_file, dpi=300, bbox_inches='tight')
        plt.close()
        
        print(f"Word cloud saved as {output_file}")
        
    except ImportError:
        print("WordCloud package not installed. Install with: pip install wordcloud")

def analyze_organization_topics(file_path, output_file):
    """Analyze which organizations focus on which topics"""
    # Read the Excel file
    df = pd.read_excel(file_path)
    
    # Clean and filter data
    df = df[(df['organization'].notna()) & (df['organization'] != 'N/A') & 
            (df['topic'].notna()) & (df['topic'] != 'N/A')]
    
    # Convert to strings and strip whitespace
    df['organization'] = df['organization'].astype(str).str.strip()
    df['topic'] = df['topic'].astype(str).str.strip()
    
    # Count organization-topic pairs
    org_topic_counts = df.groupby(['organization', 'topic']).size().reset_index(name='count')
    
    # Get top 10 organizations by sticker count
    top_orgs = df['organization'].value_counts().nlargest(10).index.tolist()
    
    # Filter for top organizations
    filtered_counts = org_topic_counts[org_topic_counts['organization'].isin(top_orgs)]
    
    # Create a pivot table
    pivot_table = filtered_counts.pivot_table(
        index='organization', 
        columns='topic', 
        values='count',
        fill_value=0
    )
    
    # Plot heatmap
    plt.figure(figsize=(14, 10))
    sns.heatmap(
        pivot_table, 
        annot=True, 
        cmap='viridis', 
        fmt='g',
        cbar_kws={'label': 'Number of Stickers'}
    )
    
    plt.title('Topics by Organization (Top 10 Organizations)', fontsize=16)
    plt.tight_layout()
    
    # Save the heatmap
    plt.savefig(output_file, dpi=300, bbox_inches='tight')
    plt.close()
    
    print(f"Organization-topic heatmap saved as {output_file}")

if __name__ == "__main__":
    # Create directories for data and output
    data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
    
    # Create directories if they don't exist
    os.makedirs(data_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    
    # File path
    file_path = os.path.join(data_dir, 'general.xlsx')
    
    print(f"Looking for data file in: {file_path}")
    
    # Define output file paths
    bar_chart_path = os.path.join(output_dir, 'topic_bar_chart.png')
    pie_chart_path = os.path.join(output_dir, 'topic_pie_chart.png')
    wordcloud_path = os.path.join(output_dir, 'topic_wordcloud.png')
    heatmap_path = os.path.join(output_dir, 'organization_topic_heatmap.png')
    
    # Analyze topic frequencies
    topic_df = analyze_topic_frequencies(file_path)
    
    print(f"Found {len(topic_df)} unique topics")
    print("Top 5 topics:")
    for i, row in topic_df.head(5).iterrows():
        print(f"  {row['Topic']}: {row['Count']} mentions")
    
    # Create visualizations
    create_bar_chart(topic_df, bar_chart_path)
    create_pie_chart(topic_df, pie_chart_path)
    create_wordcloud(topic_df, wordcloud_path)
    
    # Analyze organization-topic relationships
    analyze_organization_topics(file_path, heatmap_path)
    
    print("\nVisualization complete!")
    print("Output files:")
    print(f"1. {bar_chart_path} - Bar chart of topic frequencies")
    print(f"2. {pie_chart_path} - Pie chart of topic distribution")
    print(f"3. {wordcloud_path} - Word cloud visualization of topics (requires wordcloud package)")
    print(f"4. {heatmap_path} - Heatmap showing which organizations focus on which topics")