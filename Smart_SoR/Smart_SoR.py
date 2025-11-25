import pandas as pd
import re
import streamlit as st

# Page configuration
st.set_page_config(
    page_title="Store Performance Q&A Bot",
    page_icon="📊",
    layout="wide"
)

# Load dataset
@st.cache_data
def load_data():
    df = pd.read_csv(r"SOR_Test.csv")
    df.columns = df.columns.str.strip()
    return df

df = load_data()

def answer_question(question):
    question = question.lower()
    
    # Extract potential filters
    months = ['january', 'february', 'march', 'april', 'may', 'june', 
              'july', 'august', 'september', 'october', 'november', 'december']
    
    # Create a copy to filter
    filtered = df.copy()
    
    # Filter by month if mentioned
    month_found = None
    for month in months:
        if month in question:
            filtered = filtered[filtered['Month'].str.lower() == month]
            month_found = month
            break
    
    # Filter by year if mentioned
    year_match = re.search(r'\b(20\d{2})\b', question)
    if year_match:
        year = int(year_match.group(1))
        filtered = filtered[filtered['Year'] == year]
    
    # Filter by store if mentioned (but not for "what store" questions)
    store_found = None
    if 'what store' not in question and 'which store' not in question:
        for store in df['Store Name'].unique():
            if store.lower() in question:
                filtered = filtered[filtered['Store Name'] == store]
                store_found = store
                break
    
    # Check if asking for store identification
    asking_for_store = any(phrase in question for phrase in ['what store', 'which store', 'what location'])
    
    # Determine operation - treat "total" as "sum"
    operation = 'sum'
    if any(word in question for word in ['average', 'avg', 'mean']):
        operation = 'mean'
    elif any(word in question for word in ['maximum', 'max', 'highest']):
        operation = 'max'
    elif any(word in question for word in ['minimum', 'min', 'lowest']):
        operation = 'min'
    elif any(word in question for word in ['count', 'how many']):
        operation = 'count'
    
    # Enhanced column matching - prioritize exact matches
    column_map = {
        'new retail units': 'New Retail Units',
        'new base retail pvr': 'New Base Retail PVR',
        'new mfg incentive': 'New Mfg Incentive PVR',
        'new core pvr': 'New Core PVR',
        'new ovi pvr': 'New OVI PVR',
        'new cfs pvr': 'New CFS PVR',
        'new total deal gross pvr': 'New Total Deal Gross PVR',
        'new variable gross': 'New Variable Gross', 
        'new sales efficiency': 'Sales Efficiency',
        'used retail units': 'Used Retail Units',
        'used base retail pvr': 'Used Base Retail PVR',
        'used mfg incentive': 'Used Mfg Incentive PVR',
        'used core pvr': 'Used Core PVR',
        'used ovi pvr': 'Used OVI PVR',
        'used cfs pvr': 'Used CFS PVR',
        'used total deal gross pvr': 'Used Total Deal Gross PVR', 
        'used total deal gross pvr': 'Used Total Deal Gross PVR',
        'used variable gross': 'Used Variable Gross',
    }
    
    matched_column = None
    
    # First try exact phrase matching (prioritize longer matches)
    sorted_keys = sorted(column_map.keys(), key=len, reverse=True)
    for key in sorted_keys:
        if key in question:
            matched_column = column_map[key]
            break
    
    # If no match found, try partial matching with actual columns
    if not matched_column:
        for col in df.columns:
            # Skip Month, Year, Store Name columns
            if col in ['Month', 'Year', 'Store Name']:
                continue
            col_words = col.lower().split()
            question_words = question.split()
            # Check if significant words from column appear in question
            if any(word in question_words for word in col_words if len(word) > 4):
                matched_column = col
                break
    
    if matched_column and matched_column in filtered.columns:
        # If asking for which store has max/min
        if asking_for_store and operation in ['max', 'min']:
            if operation == 'max':
                idx = filtered[matched_column].idxmax()
            else:  # min
                idx = filtered[matched_column].idxmin()
            
            store_name = filtered.loc[idx, 'Store Name']
            value = filtered.loc[idx, matched_column]
            
            # Build descriptive response
            filters = []
            if month_found:
                filters.append(month_found.capitalize())
            if year_match:
                filters.append(str(year_match.group(1)))
            
            filter_str = " in " + " ".join(filters) if filters else ""
            return f"{store_name} has the {operation} {matched_column}{filter_str}: {value:,.2f}"
        
        # Regular operation
        if operation == 'sum':
            result = filtered[matched_column].sum()
        elif operation == 'mean':
            result = filtered[matched_column].mean()
        elif operation == 'max':
            result = filtered[matched_column].max()
        elif operation == 'min':
            result = filtered[matched_column].min()
        elif operation == 'count':
            result = filtered[matched_column].count()
        
        # Build descriptive response
        filters = []
        if month_found:
            filters.append(month_found.capitalize())
        if year_match:
            filters.append(str(year_match.group(1)))
        if store_found:
            filters.append(store_found)
        
        filter_str = " - ".join(filters) if filters else "all data"
        return f"{operation.capitalize()} of {matched_column} for {filter_str}: {result:,.2f}"
    
    return "Sorry, I couldn't understand that question. Try asking about specific columns like 'New Retail Units' or 'Used Total Deal Gross PVR'."

# Streamlit UI
st.title("📊 Store Performance Q&A Bot")
st.markdown("Ask questions about store data from January to October 2025")

# Sidebar with preset questions
st.sidebar.header("Quick Questions")
st.sidebar.markdown("Click a button to run a preset question:")

preset_questions = {
    "📈 Total New Retail Units (Oct 2025)": "What is the total New Retail Units for October 2025?",
    "🏆 Top Store - New Base Retail (Sep)": "What store has the maximum New Base Retail PVR in September?",
    "📊 Avg Used Variable Gross": "What is the average Used Variable Gross?",
    "🚗 Total Used Retail Units (2025)": "What is the total Used Retail Units for 2025?",
    "💰 Max New Total Deal Gross (Oct)": "What is the maximum New Total Deal Gross PVR in October 2025?",
    "📉 Min Used Core PVR": "What is the minimum Used Core PVR?",
}

# Initialize session state for storing questions
if 'current_question' not in st.session_state:
    st.session_state.current_question = ""

# Preset question buttons in sidebar
for label, question in preset_questions.items():
    if st.sidebar.button(label, use_container_width=True):
        st.session_state.current_question = question

st.sidebar.markdown("---")
st.sidebar.markdown("### 💡 Example Questions")
st.sidebar.markdown("""
- What is the total New Retail Units for October 2025?
- What store has the maximum New Base Retail PVR?
- What is the average Used Variable Gross for BMW of Carlsbad?
""")

# Main input area
col1, col2 = st.columns([3, 1])

with col1:
    user_question = st.text_input(
        "Ask your question:",
        value=st.session_state.current_question,
        placeholder="e.g., What is the total New Retail Units for October 2025?",
        key="question_input"
    )

with col2:
    st.markdown("<br>", unsafe_allow_html=True)
    ask_button = st.button("Ask", type="primary", use_container_width=True)

# Process question
if ask_button and user_question:
    with st.spinner("Analyzing..."):
        answer = answer_question(user_question)
        
        # Display answer in a nice format
        st.markdown("### Answer:")
        if "Sorry" in answer:
            st.warning(answer)
        else:
            st.success(answer)
        
        # Clear the preset question after use
        st.session_state.current_question = ""

# Display data preview
with st.expander("📋 View Data Preview"):
    st.dataframe(df.head(10), use_container_width=True)

# Display available stores
with st.expander("🏪 Available Stores"):
    stores = sorted(df['Store Name'].unique())
    st.write(", ".join(stores))

# Display available months
with st.expander("📅 Available Months"):
    months = sorted(df['Month'].unique())
    st.write(", ".join(months))





# def main():
#     print("="*60)
#     print("Welcome to the Store Performance Q&A Bot!")
#     print("="*60)
#     print("\nI can answer questions about your store data (Jan-Oct 2025)")
#     print("\nExample questions:")
#     print("- What is the total New Retail Units for October 2025?")
#     print("- What store has the maximum New Base Retail PVR in September?")
#     print("- What is the average Used Variable Gross for BMW of Carlsbad?")
#     print("\nType 'exit', 'quit', or 'I have no further questions' to end.")
#     print("="*60)
    
#     while True:
#         user_input = input("\nYour question: ").strip()
        
#         # Check for exit conditions
#         exit_phrases = ['exit', 'quit', 'no further questions', 'i have no further questions', 'done', 'bye']
#         if any(phrase in user_input.lower() for phrase in exit_phrases):
#             print("\nThank you for using the Store Performance Q&A Bot. Goodbye!")
#             break
        
#         # Skip empty inputs
#         if not user_input:
#             print("Please ask a question.")
#             continue
        
#         # Get and display answer
#         answer = answer_question(user_input)
#         print(f"\n{answer}")

# if __name__ == "__main__":
#     main()