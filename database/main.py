# main.py
import os
import time
import traceback
import ReviewAnalysis
import preprocess
# import temp_preprocess
def main():
    """
    Main function to orchestrate the review analysis process.
    """
    try:
        # 1. Set up API Key from environment Variable
        api_key = "AIzaSyAxk2Wog2ylp7wuQgTGdQCakzJXMoRHzO8"
        # if not api_key:
        #     raise ValueError("GOOGLE_API_KEY environment variable not set.")

        # 2. Call ReviewAnalysis to tag all non-tagged reviews with sentiment analysis

        # print("Starting Review Analysis...")
        # ReviewAnalysis.process_reviews_in_db(api_key)  # Tag reviews in the database
        # print("Review Analysis Complete.")
        # time.sleep(5) # Add some wait time to avoid hitting rate limits

        # 3. Call Preprocess to generate and populate table for further reporting.
        print("Starting Preprocessing...")

        # Set the month to process (This is where you can adjust the latest month)
        month_to_process = 2  # Change this to the latest month you want to process

        preprocess.process_reviews_and_store_data(api_key, month_to_process=month_to_process)  # aggregate and store data in new table
        print("Preprocessing Complete.")

        print("All operations completed successfully.")

    except Exception as e:
        print(f"An error occurred during the process: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    main()