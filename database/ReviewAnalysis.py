import mysql.connector
import time
import os
import json
import traceback
from google.api_core.exceptions import ResourceExhausted
import google.generativeai as genai
from datetime import datetime, date
from dateutil.relativedelta import relativedelta

categories = [
    "Cleanliness", "Menu Variety", "Portion Size", "Staff Friendliness", "Overall Experience",
    "Ambiance", "Speed of Service", "Service", "Value for Money", "Food Quality"
]

# Database configuration
db_config = {
    'user': 'root',
    'password': 'Yaswanth123.',
    'host': 'localhost',
    'database': 'genai',
    'raise_on_warnings': True
}


def generate_content_from_file(review, api_key):  # Pass api_key as argument
    """Generates sentiment and extracts information from a review using Generative AI model."""
    prompt = f"""You are an expert in analyzing customer reviews for restaurants.  For the following review, please provide the overall sentiment of the review. Also, identify any staff names mentioned along with their sentiment in the review, any dish names mentioned along with their sentiment, and identify sentiment for each matching category from the following list: {', '.join(categories)}.  Provide your response in a JSON format with the following structure:

{{
  "review_sentiment": "positive" or "negative" or "neutral",
  "dish_sentiment": {{"dish_name1": "positive/negative/neutral", "dish_name2": "positive/negative/neutral"}} or {{}},
  "staff_sentiment": {{"staff_name1": "positive/negative/neutral", "staff_name2": "positive/negative/neutral"}} or {{}},
  "category_sentiment": {{"category1": "positive/negative/neutral", "category2": "positive/negative/neutral"}} or {{}}
}}

If a field cannot be determined, set its value to an empty dictionary (for dish_sentiment, staff_sentiment, category_sentiment) or neutral for review_sentiment. Make sure the keys are always enclosed in double quotes.

Here is the review: {review}"""
    max_retries = 5
    for attempt in range(max_retries):
        try:
            genai.configure(api_key=api_key)  # Configure Gemini API *within* the function
            model = genai.GenerativeModel('gemini-2.0-flash')
            response = model.generate_content(prompt)
            return response.text.strip()
        except ResourceExhausted as e:
            if attempt < max_retries - 1:
                sleep_time = 9 ** attempt  # Exponential backoff
                print(f"Quota exceeded. Retrying in {sleep_time} seconds...")
                time.sleep(sleep_time)
            else:
                raise e
        except Exception as e:
            print(f"Error during API call: {e}")
            return None  # or raise, depending on your desired behavior



def process_reviews_in_db(api_key):  # Pass api_key to process_reviews_in_db
    """Reads reviews from the database, analyzes them, and updates the table."""

    try:
        cnx = mysql.connector.connect(**db_config)
        cursor = cnx.cursor()

        # SQL query to select rows where review_sentiment is NULL (or empty)
        select_reviews_query = "SELECT `Outlet`, `review_month`, `Year`, reviews FROM reviews_trend_dummy WHERE `review_month` = 2"  # Limit to 1000 for safety

        cursor.execute(select_reviews_query)
        review_rows = cursor.fetchall()

        if not review_rows:
            print("No new reviews to process.")
            return

        for outlet, source_month, year, review_text in review_rows:
            if review_text:
                try:
                    api_response = generate_content_from_file(review_text, api_key)  # Pass api_key

                    if api_response:
                        # Clean up the API response
                        api_response = api_response.replace("```json", "").replace("```", "").strip()

                        try:
                            data = json.loads(api_response)
                            review_sentiment = data.get('review_sentiment', 'neutral')
                            dish_sentiment = data.get('dish_sentiment', {})
                            staff_sentiment = data.get('staff_sentiment', {})
                            category_sentiment = data.get('category_sentiment', {})

                            # Prepare the SQL query to update the row
                            update_review_query = """
                                UPDATE reviews_trend_dummy
                                SET review_sentiment = %s, dish_sentiment = %s, staff_sentiment = %s, category_sentiment = %s
                                WHERE outlet = %s AND `review_month` = %s AND `Year` = %s AND reviews = %s
                            """

                            # Convert dictionaries to JSON strings
                            data_to_update = (
                                review_sentiment,
                                json.dumps(dish_sentiment),
                                json.dumps(staff_sentiment),
                                json.dumps(category_sentiment),
                                outlet,
                                source_month,
                                year,
                                review_text
                            )

                            # Execute the update query
                            cursor.execute(update_review_query, data_to_update)
                            cnx.commit()

                            print(f"Review for outlet {outlet}, month {source_month}, year {year} updated successfully.")

                        except json.JSONDecodeError as e:
                            print(f"Error decoding JSON response for outlet {outlet}, month {source_month}, year {year}: {e}\nResponse was: {api_response}")
                            traceback.print_exc()
                            with open("json_error_log.txt", "a") as f:
                                f.write(f"Outlet: {outlet}, Month: {source_month}, Year: {year}\n")
                                f.write(f"Response: {api_response}\n")
                                f.write(traceback.format_exc() + "\n")

                        except Exception as e:
                            print(f"Error during database update: {e}")
                            cnx.rollback()

                    else:
                        print(f"No response from API for outlet {outlet}, month {source_month}, year {year}")

                except Exception as e:
                    print(f"Error processing review for outlet {outlet}, month {source_month}, year {year}: {e}")
                    cnx.rollback()

            else:
                print(f"No review text found for outlet {outlet}, month {source_month}, year {year}. Skipping...")

        print("All reviews processed and updated in the database.")

    except mysql.connector.Error as err:
        print(f"Error connecting to MySQL: {err}")

    finally:
        if cursor:
            cursor.close()
        if cnx:
            cnx.close()


# Example usage
# if __name__ == "__main__":
#     api_key = os.environ.get("GOOGLE_API_KEY")
#     process_reviews_in_db(api_key)