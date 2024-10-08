import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import glob

# Title for the Streamlit app
st.title('Ticket Data Analysis')

# Function to load data from uploaded files
@st.cache_data
def load_data(files):
    li = []
    for file in files:
        df = pd.read_excel(file, '-', index_col=None, header=0)
        li.append(df)
    df = pd.concat(li, axis=0, ignore_index=True)
    return df.drop_duplicates()

# Upload Excel files
uploaded_files = st.file_uploader("Upload Excel files", accept_multiple_files=True, type=['xls', 'xlsx'])

if uploaded_files:
    # Load all uploaded Excel files
    df = load_data(uploaded_files)

    # Function to apply custom filtering logic
    def filter_tickets(group):
        if 'CLOSED' in group['Status'].values:
            return group[group['Status'] == 'CLOSED']
        elif 'RESOLVED' in group['Status'].values:
            return group[group['Status'] == 'RESOLVED']
        elif 'OPEN' in group['Status'].values:
            return group[group['Status'] == 'OPEN']

    # Group by 'Reference' and apply the filtering logic
    df = df.groupby('Reference').apply(filter_tickets).reset_index(drop=True)

    # Sort by 'Date/Time Logged'
    df = df.sort_values(by='Date/Time Logged')

    # Convert 'Date/Time Logged' and 'Date Resolved' to datetime
    df['Date Logged'] = pd.to_datetime(df['Date/Time Logged']).dt.date
    df['Date Resolved'] = pd.to_datetime(df['Date Resolved']).dt.date

    # Count tickets by logged date
    df_count = df['Date Logged'].value_counts().reset_index()
    df_count.columns = ['Date', 'Ticket Count']
    df_count = df_count.sort_values('Date')

    # Count tickets by resolved date
    df_resolved_count = df['Date Resolved'].value_counts().reset_index()
    df_resolved_count.columns = ['Date', 'Ticket Count']
    df_resolved_count = df_resolved_count.sort_values('Date')

    # Create bar chart for tickets logged and resolved
    st.subheader('Bar Chart')
    bar_fig = go.Figure()

    # Add traces for logged and resolved tickets
    bar_fig.add_trace(go.Bar(x=df_count['Date'], y=df_count['Ticket Count'], name='Logged Tickets'))
    bar_fig.add_trace(go.Bar(x=df_resolved_count['Date'], y=df_resolved_count['Ticket Count'], name='Resolved Tickets'))

    # Update layout to include date range slider
    bar_fig.update_layout(
        title='Number of Tickets and Tickets Resolved per Day',
        xaxis_title='Date',
        yaxis_title='Ticket Count',
        xaxis=dict(
            rangeselector=dict(
                buttons=list([
                    dict(count=1, label='1m', step='month', stepmode='backward'),
                    dict(count=3, label='3m', step='month', stepmode='backward'),
                    dict(count=6, label='6m', step='month', stepmode='backward'),
                    dict(count=1, label='YTD', step='year', stepmode='todate'),
                    dict(count=1, label='1y', step='year', stepmode='backward'),
                    dict(step='all')
                ])
            ),
            rangeslider=dict(visible=True),
            type='date'
        ),
        barmode='group'
    )

    # Display bar chart
    st.plotly_chart(bar_fig)

    # Create line chart for tickets logged and resolved
    st.subheader('Line Chart')
    line_fig = go.Figure()

    # Add traces for tickets logged and resolved
    line_fig.add_trace(go.Scatter(x=df_count['Date'], y=df_count['Ticket Count'], mode='lines', name='Logged Tickets'))
    line_fig.add_trace(go.Scatter(x=df_resolved_count['Date'], y=df_resolved_count['Ticket Count'], mode='lines', name='Resolved Tickets'))

    # Update layout to include date range slider
    line_fig.update_layout(
        title='Number of Tickets Logged and Resolved per Day',
        xaxis_title='Date',
        yaxis_title='Ticket Count',
        xaxis=dict(
            rangeselector=dict(
                buttons=list([
                    dict(count=1, label='1m', step='month', stepmode='backward'),
                    dict(count=3, label='3m', step='month', stepmode='backward'),
                    dict(count=6, label='6m', step='month', stepmode='backward'),
                    dict(count=1, label='YTD', step='year', stepmode='todate'),
                    dict(count=1, label='1y', step='year', stepmode='backward'),
                    dict(step='all')
                ])
            ),
            rangeslider=dict(visible=True),
            type='date'
        )
    )

    # Display line chart
    st.plotly_chart(line_fig)

else:
    st.info("Please upload one or more Excel files to analyze.")
