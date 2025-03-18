import os
import time
import traceback
import ReviewAnalysis
import preprocess
import Formatting

def hello_http(request):
    """
    HTTP Cloud Function entry point for review analysis.
    Args:
        request (flask.Request): The request object.
        <https://flask.palletsprojects.com/en/2.3.x/api/#flask.Request>
    Returns:
        The response text, or any other valid response.
    """
    try:
        # 1. Get API Key (Securely!  Do NOT hardcode in production)
        # api_key = os.environ.get('API_KEY')  # Retrieve from environment variables
        # if not api_key:
        #     return "Error: API_KEY environment variable not set.", 500  # Indicate missing key

        api_key = "AIzaSyAxk2Wog2ylp7wuQgTGdQCakzJXMoRHzO8"
        month_to_process = int(request.args.get('month', 2)) 
        # 2. Review Analysis
        print("Starting Review Analysis...")
        ReviewAnalysis.process_reviews_in_db(api_key, month_to_process=month_to_process)
        print("Review Analysis Complete.")
        time.sleep(5)

        # 3. Preprocessing
        print("Starting Preprocessing...")
        # Get month from query parameter, default to 2
        preprocess.process_reviews_and_store_data(api_key, month_to_process=month_to_process)
        print("Preprocessing Complete.")


        # 4. Formatting (Adjust to handle optional parameters from HTTP request)
        print("Starting Formatting...")
        outlet_value = request.args.get('outlet', "South Plainfield")  # Get outlet from query parameter
        review_month_value = month_to_process
        num_months = int(request.args.get('num_months', 3)) #Get number of months from query parameter, default to 3

        data = Formatting.fetch_data_from_db(outlet_value, review_month_value, num_months=num_months)
        path = f"A2B_{outlet_value}_{review_month_value}.docx"
        if data is not None:
            Formatting.create_word_document(outlet_value, review_month_value, data, output_filename=path)
            #Cloud functions cannot directly open files, comment out the following line
            #Formatting.open_word_file(path) #Cloud functions cannot directly open files
        else:
            return "Failed to fetch data. Check credentials and query.", 500

        print("Formatting Complete.")
        return "All operations completed successfully."

    except Exception as e:
        print(f"An error occurred: {e}")
        traceback.print_exc()
        return f"An error occurred: {e}", 500 # Return error to the HTTP client