import pandas as pd
import streamlit as st
import plotly.express as px

# Load the Excel file
file_path = r"C:\Users\Avni.Tonger\Downloads\sentimentanalysis4.xlsx"
df = pd.read_excel(file_path, engine='openpyxl')

# Display columns
df_display = df[['Number', 'Description', 'Sentiment_Score', 'Urgency', 'Category', 'Priority']]

# Create a pie chart for sentiment score distribution
sentiment_bins = [0, 4, 7, 10]
sentiment_labels = ['Low (0-4)', 'Medium (5-7)', 'High (8-10)']
df['Sentiment_Category'] = pd.cut(df['Sentiment_Score'], bins=sentiment_bins, labels=sentiment_labels, include_lowest=True)
sentiment_counts = df['Sentiment_Category'].value_counts().reindex(sentiment_labels)

fig_sentiment_pie = px.pie(
    values=sentiment_counts,
    names=sentiment_counts.index,
    title='Sentiment Score Distribution',
    color=sentiment_counts.index,
    color_discrete_map={'Low (0-4)': '#FFD700', 'Medium (5-7)': '#FFA500', 'High (8-10)': '#FF4500'},
    hole=0.3
)
fig_sentiment_pie.update_layout(
    title_font=dict(size=24, color='#333'),
    legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
    margin=dict(l=20, r=20, t=40, b=20),
    paper_bgcolor='#f9f9f9',
    plot_bgcolor='#f9f9f9'
)
fig_sentiment_pie.update_traces(textinfo='percent+label')

# Include Sentiment_Category in df_display
df_display['Sentiment_Category'] = df['Sentiment_Category']

# Sort incidents based on Sentiment_Score and Urgency
sorted_df_display = df_display.sort_values(by=['Sentiment_Score', 'Urgency'], ascending=[False, False])

# Create a bar chart for priority distribution
fig_priority_bar = px.bar(
    sorted_df_display,
    x='Priority',
    color='Sentiment_Category',
    title='Priority Distribution by Sentiment Score',
    color_discrete_map={'Low (0-4)': '#FFD700', 'Medium (5-7)': '#FFA500', 'High (8-10)': '#FF4500'}
)
fig_priority_bar.update_layout(
    title_font=dict(size=24, color='#333'),
    legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
    margin=dict(l=20, r=20, t=40, b=20),
    paper_bgcolor='#f9f9f9',
    plot_bgcolor='#f9f9f9'
)

# Streamlit layout
st.title("Incident Sentiment Analysis Dashboard")
st.plotly_chart(fig_sentiment_pie)

st.plotly_chart(fig_priority_bar)

st.subheader("Incidents to Address Immediately")
st.dataframe(sorted_df_display)