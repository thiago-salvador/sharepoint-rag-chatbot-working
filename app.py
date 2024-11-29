import os
import streamlit as st
from dotenv import load_dotenv
import time
from typing import List, Dict
from utils.sharepoint import SharePointConnector
from utils.indexer import DocumentIndexer
from utils.chat_engine import ChatEngine

# Load environment variables
load_dotenv()

# Configure Streamlit page
st.set_page_config(
    page_title="SharePoint RAG Chatbot",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize session state
if 'messages' not in st.session_state:
    st.session_state.messages = []
if 'documents' not in st.session_state:
    st.session_state.documents = []
if 'connected' not in st.session_state:
    st.session_state.connected = False

def initialize_sharepoint():
    """Initialize SharePoint connection and index documents"""
    try:
        # Create SharePoint connector
        connector = SharePointConnector(
            st.session_state.sharepoint_url,
            st.session_state.site_name,
            st.session_state.username,
            st.session_state.password
        )
        
        # Get documents
        with st.spinner('Fetching documents from SharePoint...'):
            documents = connector.get_all_documents()
            st.session_state.documents = documents
        
        # Create document indexer
        indexer = DocumentIndexer()
        
        # Index documents
        with st.spinner('Indexing documents...'):
            index = indexer.index_documents(documents)
            st.session_state.index = index
        
        # Initialize chat engine
        st.session_state.chat_engine = ChatEngine(index)
        st.session_state.connected = True
        
        # Show success message
        st.success(f'Successfully connected and indexed {len(documents)} documents!')
        
    except Exception as e:
        st.error(f'Error: {str(e)}')
        st.session_state.connected = False

# Sidebar for configuration
with st.sidebar:
    st.title('Configuration')
    
    # SharePoint credentials input
    with st.form('sharepoint_credentials'):
        st.subheader('SharePoint Credentials')
        
        # Use environment variables as defaults
        sharepoint_url = st.text_input(
            'SharePoint URL',
            value=os.getenv('SHAREPOINT_URL', ''),
            type='default'
        )
        site_name = st.text_input(
            'Site Name',
            value=os.getenv('SHAREPOINT_SITE_NAME', ''),
            type='default'
        )
        username = st.text_input(
            'Username',
            value=os.getenv('SHAREPOINT_USERNAME', ''),
            type='default'
        )
        password = st.text_input(
            'Password',
            value=os.getenv('SHAREPOINT_PASSWORD', ''),
            type='password'
        )
        
        # Store values in session state
        if sharepoint_url:
            st.session_state.sharepoint_url = sharepoint_url
        if site_name:
            st.session_state.site_name = site_name
        if username:
            st.session_state.username = username
        if password:
            st.session_state.password = password
        
        connect_button = st.form_submit_button('Connect to SharePoint')
        
        if connect_button:
            initialize_sharepoint()
    
    # Show document list when connected
    if st.session_state.connected and st.session_state.documents:
        st.subheader('Indexed Documents')
        for doc in st.session_state.documents:
            st.markdown(f"- {doc['name']}")

# Main chat interface
st.title('SharePoint RAG Chatbot')

# Display chat messages
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])
        if "sources" in message:
            st.markdown("Sources:")
            for source in message["sources"]:
                st.markdown(f"- {source}")

# Chat input
if prompt := st.chat_input(
    "Ask a question about your SharePoint documents",
    disabled=not st.session_state.connected
):
    # Add user message to chat history
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)
    
    # Generate response
    with st.chat_message("assistant"):
        if not st.session_state.connected:
            st.error('Please connect to SharePoint first!')
        else:
            try:
                with st.spinner('Thinking...'):
                    response, sources = st.session_state.chat_engine.get_response(
                        prompt,
                        st.session_state.messages
                    )
                    
                    st.markdown(response)
                    
                    if sources:
                        st.markdown("Sources:")
                        for source in sources:
                            st.markdown(f"- {source}")
                    
                    # Add assistant response to chat history
                    st.session_state.messages.append({
                        "role": "assistant",
                        "content": response,
                        "sources": sources
                    })
            
            except Exception as e:
                st.error(f'Error generating response: {str(e)}')

# Instructions
if not st.session_state.connected:
    st.info("""
    ### Getting Started
    1. Enter your SharePoint credentials in the sidebar
    2. Click 'Connect to SharePoint' to establish connection
    3. Wait for documents to be indexed
    4. Start asking questions about your documents
    """)
