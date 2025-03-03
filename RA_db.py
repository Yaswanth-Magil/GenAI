import time
import json
import traceback
import mysql.connector  # For MySQL connection
from google.api_core.exceptions import ResourceExhausted
import google.generativeai as genai

categories = [
    "Cleanliness", "Menu Variety", "Portion Size", "Staff Friendliness", "Overall Experience",
    "Ambiance", "Speed of Service", "Service", "Value for Money", "Food Quality"
]

def generate_content_from_file(review):
    """Generates sentiment and extracts information from a review using Generative AI model."""
    prompt = f"""You are an expert in analyzing customer reviews for restaurants. For the following review, please identify the sentiment (positive, negative, or neutral), any staff names mentioned, any dish names mentioned, and the *single most relevant* category from this list: {', '.join(categories)}.  Provide your response in a JSON format with the following structure:

{{
  "sentiment": "positive" or "negative" or "neutral",
  "staff_names": ["list", "of", "staff", "names"] or [],
  "dish_names": ["list", "of", "dish", "names"] or [],
  "category": "one of the categories from the list" or null
}}

If a field cannot be determined, set its value to null (for category) or an empty list (for staff_names and dish_names).  Make sure the keys are always enclosed in double quotes.

Here is the review: {review}"""
    max_retries = 5
    for attempt in range(max_retries):
        try:
            response = genai.GenerativeModel('gemini-2.0-flash').generate_content(prompt)
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

def get_reviews_from_db(connection):
    """Fetches reviews from the database."""
    query = "SELECT Review FROM reviews WHERE Review IS NOT NULL"
    cursor = connection.cursor()
    cursor.execute(query)
    reviews = cursor.fetchall()
    cursor.close()
    return reviews

def update_review_in_db(connection, review, sentiment, staff_names, food_names, category):
    """Updates the review's analysis results back in the database."""
    query = """
    UPDATE reviews
    SET Sentiment = %s, staff_name = %s, Food_name = %s, Category = %s
    WHERE Review = %s
    """
    cursor = connection.cursor()
    cursor.execute(query, (sentiment, ', '.join(staff_names), ', '.join(food_names), category, review))
    connection.commit()
    cursor.close()

def process_reviews(connection):
    """Processes reviews from the database and adds sentiment and extractions."""
    reviews = get_reviews_from_db(connection)

    for review in reviews:
        review_text = review[0]
        if review_text:
            try:
                api_response = generate_content_from_file(review_text)

                if api_response:
                    print(f"API Response: {api_response}")  # For debugging

                    # Remove extra characters before and after the JSON
                    api_response = api_response.replace("```json", "").replace("```", "").strip()

                    try:
                        data = json.loads(api_response)
                        sentiment = data.get('sentiment', 'Unknown')
                        staff_names = data.get('staff_names', [])
                        food_names = data.get('dish_names', [])
                        category = data.get('category', 'Unknown')

                        update_review_in_db(connection, review_text, sentiment, staff_names, food_names, category)

                        print(f"Review: {review_text}\nSentiment: {sentiment}\nStaff: {staff_names}\nDishes: {food_names}\nCategory: {category}\n")

                    except json.JSONDecodeError as e:
                        print(f"Error decoding JSON response for review {review_text}: {e}\nResponse was: {api_response}")
                        traceback.print_exc()
                        with open("json_error_log.txt", "a") as f:
                            f.write(f"Review: {review_text}\n")
                            f.write(f"Response: {api_response}\n")
                            f.write(traceback.format_exc() + "\n")
                        
                        # Update the review with error status
                        update_review_in_db(connection, review_text, "JSON Error", [], [], "Unknown")
                    except UnicodeDecodeError as e:
                        print(f"UnicodeDecodeError: {e}")
                        update_review_in_db(connection, review_text, "Encoding Error", [], [], "Unknown")

                else:
                    print(f"No response from API for review {review_text}")
                    update_review_in_db(connection, review_text, "API Error", [], [], "Unknown")

            except Exception as e:
                print(f"Error processing review {review_text}: {e}")
                update_review_in_db(connection, review_text, "Error", [], [], "Unknown")

        else:
            print("No review text found. Skipping...\n")
            continue

def main():
    """Main function to execute the sentiment analysis."""
    start_time = time.time()  # Start time

    api_key = "AIzaSyAxk2Wog2ylp7wuQgTGdQCakzJXMoRHzO8"
    genai.configure(api_key=api_key)

    try:
        # Set up the MySQL database connection
        connection = mysql.connector.connect(
            host="localhost",
            database="genai",
            user="root",
            password="Yaswanth123."
        )

        process_reviews(connection)

    except Exception as e:
        print(f"Error connecting to the database: {e}")

    finally:
        if connection:
            connection.close()

    end_time = time.time()  # End time
    execution_time = end_time - start_time
    print(f"Total execution time: {execution_time:.2f} seconds")


if __name__ == "__main__":
    main()