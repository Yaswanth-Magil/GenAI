# main.py
import os
import ReviewAnalysis2
import preprocess
import google.generativeai as genai

def main():
    """Main function to orchestrate the review analysis and data extraction."""

    # Get API Key
    api_key = "AIzaSyAxk2Wog2ylp7wuQgTGdQCakzJXMoRHzO8"  # Get from environment variable
    
    # if not api_key:
    #     print("Error: GEMINI_API_KEY environment variable not set. Provide your API key in the environment variable")
    #     return

    # Configure Google Generative AI
    genai.configure(api_key=api_key)

    # Define file paths
    input_excel_file = "/Users/yash/Downloads/Today/Splitted/GenAI/input.xlsx"  # Path to your input Excel file
    output_excel_file = "/Users/yash/Downloads/Today/Splitted/GenAI/output.xlsx"  # Path for the aggregated output Excel file


    # Process Reviews (Sentiment Analysis)
    print("Starting sentiment analysis...")
    try:
        ReviewAnalysis2.process_reviews(input_excel_file)
        print("Sentiment analysis completed.")
    except Exception as e:
        print(f"Error during sentiment analysis: {e}")


    # Process Excel and Extract Data (Aggregation, Summarization, Competition Analysis)
    print("Starting data extraction and aggregation...")
    try:
        preprocess.process_excel_and_extract_data(input_excel_file, output_excel_file, api_key)
        print("Data extraction and aggregation completed.")
    except Exception as e:
        print(f"Error during data extraction and aggregation: {e}")


if __name__ == "__main__":
    main()