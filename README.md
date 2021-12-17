# data_processing

# link to web application deployed using Streamlit: https://share.streamlit.io/clafrance7/data_processing/main/main.py

This is a program I wrote to automate all of the data processing and formatting that
I frequently do manually in Excel at work. It reads in testing data as a csv and formats the data
in a way that can be copy and pasted into GraphPad Prism for data analysis.
It also runs basic statistics (mean, standard deviation, and %CV) on replicate tests.
Detects shifts in the search window of the test strip reader to flag potential misreads or 
strip background artifacts.
Will generate a document to archive strip images from the test strip reader.
  
