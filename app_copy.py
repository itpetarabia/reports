import streamlit as st
import time

# Text files

text_contents = '''
Foo, Bar
123, 456
789, 000
'''

# Different ways to use the API

# st.download_button('Download CSV', text_contents, 'text/csv')
# st.download_button('Download CSV', text_contents)  # Defaults to 'text/plain'
st.session_state.processed = False
if st.button('Do Something'):
    pass
    with st.spinner(text="In progress..."):
        time.sleep(5)
    # st.session_state.processed = True
    st.write('Done!')

    # if st.session_state.processed == True:
    with open('Branch_Daily_Sales_Report_Sample.xlsx', 'rb') as f:
        st.download_button('Download CSV', f, file_name='BranchReport.xlsx')  # Defaults to 'text/plain'

# ---
# Binary files

binary_contents = b'whatever'

# Different ways to use the API

# st.download_button('Download file', binary_contents)  # Defaults to 'application/octet-stream'

# with open('myfile.zip', 'rb') as f:
#    st.download_button('Download Zip', f, file_name='archive.zip')  # Defaults to 'application/octet-stream'

# You can also grab the return value of the button,
# just like with any other button.

# if st.download_button('Download CSV', text_contents):
#    st.write('Thanks for downloading!')