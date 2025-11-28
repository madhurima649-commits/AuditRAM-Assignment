#Approach and Technical Notes

1. Objective
To build a simple tool where a user uploads a PDF, and the system extracts key audit information such as Topic, Subtopic, Risk, Impact, Control, and Law/Regulation.

2. Approach Summary

a. PDF Upload (Streamlit)
I used Streamlit to create a clean web interface.
It allows the user to upload a PDF easily.

b. Text Extraction (PyPDF2)
The PDF pages are read one by one, and the text is extracted in plain format.

c. Text Processing
The extracted text is cleaned and scanned for keywords like:
Audit Topic
Subtopic
Risk Description
Business Impact
Control
Law / Regulation
Whenever a keyword is found, the next lines are saved under that category.

d. Building the Output Table
The extracted details are organized in a structured format using a Pandas DataFrame.

e. Displaying the Result
The table is shown on the Streamlit page, and the user can download it if needed.

3. Conclusion
The solution provides a simple, automated way to upload a PDF, extract important audit information, organize it into a table, and display it on a web interface.
