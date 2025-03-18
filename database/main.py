# main.py
# main.py
import os
import time
import traceback
import ReviewAnalysis
import preprocess
import Formatting
# import temp_preprocess
def main():
    """
    Main function to orchestrate the review analysis process.
    """
    try:
        # 1. Set up API Key from environment Variable
        api_key = "AIzaSyAxk2Wog2ylp7wuQgTGdQCakzJXMoRHzO8"
        
        month_to_process = 2 

        # 2. Call ReviewAnalysis to tag all non-tagged reviews with sentiment analysis

        # print("Starting Review Analysis...")
        # Change this to the latest month you want to process
        # ReviewAnalysis.process_reviews_in_db(api_key, month_to_process=month_to_process)  # Tag reviews in the database
        # print("Review Analysis Complete.")
        # time.sleep(5) # Add some wait time to avoid hitting rate limits

        # 3. Call Preprocess to generate and populate table for further reporting.
        print("Starting Preprocessing...")

        # Set the month to process (This is where you can adjust the latest month)
        # month_to_process = 2  # Change this to the latest month you want to process # removed because this has been already defined above.

        preprocess.process_reviews_and_store_data(api_key, month_to_process=month_to_process)  # aggregate and store data in new table
        print("Preprocessing Complete.")

        # 4. Call Formatting to generate the Word document.
        # print("Starting Formatting...")
        # outlet_value = "South Plainfield"  # Replace with the desired outlet value
        # review_month_value = month_to_process  # Use the month that was just processed
        # num_months = 3 # setting it to three.

        # data = Formatting.fetch_data_from_db(outlet_value, review_month_value, num_months=num_months)
        # path = f"A2B_{outlet_value}_{review_month_value}.docx"
        # if data is not None:
        #     Formatting.create_word_document(outlet_value, review_month_value, data,
        #                          output_filename=path)
        #     Formatting.open_word_file(path)
        # else:
        #     print("Failed to fetch data. Check credentials and query.")
        # print("Formatting Complete.")


        print("All operations completed successfully.")

    except Exception as e:
        print(f"An error occurred during the process: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    main()