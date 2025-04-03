import streamlit as st
import docx
import requests
from bs4 import BeautifulSoup
import re
from urllib.parse import urlparse
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import io
import pandas as pd
from collections import OrderedDict
import asyncio
import aiohttp
from concurrent.futures import ThreadPoolExecutor
import nest_asyncio
nest_asyncio.apply()

def extract_hyperlinks_from_docx(docx_file):
    """Extract all hyperlinks from a DOCX file, maintaining order and tracking duplicates."""
    doc = docx.Document(docx_file)
    urls_list = []  # Keep all URLs in order
    url_counts = {}  # Track number of occurrences
    
    # Get all hyperlinks through document relationships
    for rel in doc.part.rels.values():
        if rel.reltype == RT.HYPERLINK:
            if rel._target.startswith('http'):
                urls_list.append(rel._target)
                url_counts[rel._target] = url_counts.get(rel._target, 0) + 1
    
    # Also check for any URLs in the text (as backup)
    url_pattern = r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'
    
    # Check paragraphs
    for paragraph in doc.paragraphs:
        found_urls = re.findall(url_pattern, paragraph.text)
        for url in found_urls:
            urls_list.append(url)
            url_counts[url] = url_counts.get(url, 0) + 1
    
    # Check tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    found_urls = re.findall(url_pattern, paragraph.text)
                    for url in found_urls:
                        urls_list.append(url)
                        url_counts[url] = url_counts.get(url, 0) + 1
    
    # Get hyperlinks from headers and footers
    for section in doc.sections:
        # Header
        if section.header:
            for paragraph in section.header.paragraphs:
                found_urls = re.findall(url_pattern, paragraph.text)
                for url in found_urls:
                    urls_list.append(url)
                    url_counts[url] = url_counts.get(url, 0) + 1
        # Footer
        if section.footer:
            for paragraph in section.footer.paragraphs:
                found_urls = re.findall(url_pattern, paragraph.text)
                for url in found_urls:
                    urls_list.append(url)
                    url_counts[url] = url_counts.get(url, 0) + 1
    
    # Create a list of dictionaries with URL info
    url_info = []
    seen_urls = set()
    for url in urls_list:
        is_duplicate = url in seen_urls
        url_info.append({
            'URL': url,
            'Occurrences': url_counts[url],
            'Is Duplicate': 'Yes' if is_duplicate else 'No'
        })
        seen_urls.add(url)
    
    return url_info

async def check_url_async(session, url):
    """Asynchronously check if a URL is accessible and return its status and content."""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        async with session.get(url, headers=headers, timeout=10) as response:
            text = await response.text()
            
            # Try to get readable content using BeautifulSoup
            soup = BeautifulSoup(text, 'html.parser')
            
            # Remove script and style elements
            for script in soup(["script", "style"]):
                script.decompose()
                
            # Get text content
            content = soup.get_text()
            
            # Clean up text
            lines = (line.strip() for line in content.splitlines())
            chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
            content = ' '.join(chunk for chunk in chunks if chunk)
            
            return {
                'status': 'Success',
                'status_code': response.status,
                'content_preview': content[:500] + '...' if len(content) > 500 else content
            }
    except Exception as e:
        return {
            'status': 'Error',
            'status_code': getattr(e, 'status', None) if hasattr(e, 'status') else None,
            'content_preview': str(e)
        }

async def check_urls_batch(urls_dict):
    """Check multiple URLs concurrently in batches."""
    async with aiohttp.ClientSession() as session:
        tasks = []
        results = {'success': [], 'failed': []}
        
        # Create tasks for all URLs
        for url, occurrences in urls_dict.items():
            task = asyncio.ensure_future(check_url_async(session, url))
            tasks.append((url, occurrences, task))
        
        # Process URLs in batches of 5
        batch_size = 5
        for i in range(0, len(tasks), batch_size):
            batch = tasks[i:i + batch_size]
            
            # Wait for the current batch to complete
            for url, occurrences, task in batch:
                result = await task
                result['occurrences'] = occurrences
                
                if result['status'] == 'Success':
                    results['success'].append({'url': url, **result})
                else:
                    results['failed'].append({'url': url, **result})
                
                # Update progress (if needed)
                progress = (i + len(batch)) / len(tasks)
                if 'progress_bar' in st.session_state:
                    st.session_state.progress_bar.progress(progress)
                if 'status_text' in st.session_state:
                    st.session_state.status_text.text(f"Processed {i + len(batch)} of {len(tasks)} URLs")
        
        return results

def main():
    st.set_page_config(page_title="Document URL Checker", layout="wide")
    
    # Add custom CSS
    st.markdown("""
        <style>
        .success-url { color: #28a745; }
        .failed-url { color: #dc3545; }
        .stProgress > div > div > div > div { background-color: #28a745; }
        .duplicate-url { background-color: #e8eaed; }
        </style>
    """, unsafe_allow_html=True)
    
    st.title("📄 Document URL Checker")
    st.write("Upload a Word document (.docx) to check all URLs contained within it.")
    
    # File uploader
    uploaded_file = st.file_uploader("Choose a DOCX file", type="docx")
    
    if uploaded_file is not None:
        try:
            # Create a bytes buffer for the uploaded file
            docx_bytes = io.BytesIO(uploaded_file.read())
            
            with st.spinner('🔍 Extracting URLs from document...'):
                url_info = extract_hyperlinks_from_docx(docx_bytes)
            
            if not url_info:
                st.warning("⚠️ No URLs found in the document.")
            else:
                total_urls = len(url_info)
                unique_urls = len(set(item['URL'] for item in url_info))
                duplicate_count = total_urls - unique_urls
                
                st.success(f"✨ Found {total_urls} total URLs ({unique_urls} unique, {duplicate_count} duplicates)")
                
                # Display URLs in a table
                url_df = pd.DataFrame(url_info)
                
                # Style the dataframe to highlight duplicates
                def highlight_duplicates(row):
                    return ['background-color: #e8eaed' if row['Is Duplicate'] == 'Yes' else '' for _ in row]
                
                styled_df = url_df.style.apply(highlight_duplicates, axis=1)
                st.dataframe(styled_df, use_container_width=True)
                
                # Add a button to start checking URLs
                if st.button("🚀 Start Checking URLs"):
                    # Initialize progress bar and status
                    st.session_state.progress_bar = st.progress(0)
                    st.session_state.status_text = st.empty()
                    
                    # Create two columns for results
                    col1, col2 = st.columns(2)
                    
                    # Check URLs concurrently
                    unique_urls = {item['URL']: item['Occurrences'] for item in url_info}
                    results = asyncio.run(check_urls_batch(unique_urls))
                    
                    # Clear the status text
                    st.session_state.status_text.empty()
                    
                    # Show results in columns
                    with col1:
                        st.markdown("### ✅ Working URLs")
                        if results['success']:
                            success_data = []
                            for result in results['success']:
                                success_data.append({
                                    'URL': result['url'],
                                    'Status': result['status_code'],
                                    'Occurrences': result['occurrences'],
                                    'Content Preview': result['content_preview'][:150] + '...' if len(result['content_preview']) > 150 else result['content_preview']
                                })
                            success_df = pd.DataFrame(success_data)
                            st.dataframe(success_df, use_container_width=True)
                        else:
                            st.write("No working URLs found.")
                    
                    with col2:
                        st.markdown("### ❌ Failed URLs")
                        if results['failed']:
                            failed_data = []
                            for result in results['failed']:
                                failed_data.append({
                                    'URL': result['url'],
                                    'Status': result['status_code'],
                                    'Occurrences': result['occurrences'],
                                    'Error Message': result['content_preview'][:150] + '...' if len(result['content_preview']) > 150 else result['content_preview']
                                })
                            failed_df = pd.DataFrame(failed_data)
                            st.dataframe(failed_df, use_container_width=True)
                        else:
                            st.write("No failed URLs found.")
                    
                    # Show statistics
                    st.markdown("---")
                    st.markdown("### 📊 Statistics")
                    stats_col1, stats_col2, stats_col3, stats_col4 = st.columns(4)
                    with stats_col1:
                        st.metric("Total URLs", total_urls)
                    with stats_col2:
                        st.metric("Unique URLs", unique_urls)
                    with stats_col3:
                        st.metric("Working URLs", len(results['success']))
                    with stats_col4:
                        st.metric("Failed URLs", len(results['failed']))
                    
                    # Add export functionality
                    if st.button("📥 Export Results to CSV"):
                        # Create a DataFrame with all results
                        export_data = []
                        for result in results['success'] + results['failed']:
                            export_data.append({
                                'URL': result['url'],
                                'Status': result['status'],
                                'Status Code': result['status_code'],
                                'Times Found': result['occurrences'],
                                'Details': result['content_preview'][:200]
                            })
                        df = pd.DataFrame(export_data)
                        
                        # Convert DataFrame to CSV
                        csv = df.to_csv(index=False)
                        st.download_button(
                            label="Download CSV",
                            data=csv,
                            file_name="url_check_results.csv",
                            mime="text/csv"
                        )
                
        except Exception as e:
            st.error(f"An error occurred while processing the file: {str(e)}")
            st.exception(e)

if __name__ == "__main__":
    main() 