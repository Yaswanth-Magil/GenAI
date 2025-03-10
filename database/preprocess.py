import mysql.connector
import time
import os
import json
import traceback
from collections import defaultdict, OrderedDict
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

def aggregate_counts(counts, api_key):
    """Aggregates counts of similar dishes or staff names using Gemini's understanding."""
    aggregated_counts = defaultdict(int)
    processed = set()

    for item1 in counts:
        if item1 in processed:
            continue

        total_count = counts[item1]
        processed.add(item1)

        for item2 in counts:
            if item2 != item1 and item2 not in processed:
                similarity = check_similarity(item1, item2, api_key)

                if similarity >= 0.8:  # Adjust similarity threshold as needed
                    total_count += counts[item2]
                    processed.add(item2)
                    print(f"Aggregating '{item1}' and '{item2}' (Similarity: {similarity:.2f})")

        aggregated_counts[item1] = total_count

    return aggregated_counts


def check_similarity(item1, item2, api_key):
    """Checks the similarity between two items using Gemini."""
    prompt = f"""You are an expert in language semantics.  Determine the semantic similarity between the two phrases provided below, taking into account potential misspellings and variations in phrasing.  Return a score between 0 and 1, where 0 means completely dissimilar and 1 means essentially the same.

Phrase 1: {item1}
Phrase 2: {item2}

Provide only a number (the similarity score) with no other text."""

    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-2.0-flash')
        response = model.generate_content(prompt)
        try:
            similarity = float(response.text.strip())
            if 0 <= similarity <= 1:
                return similarity
            else:
                print(f"Warning: Similarity score out of range: {similarity}")
                return 0  # Treat out-of-range scores as dissimilar
        except ValueError:
            print(f"Error: Could not parse similarity score: {response.text}")
            return 0  # Treat unparsable responses as dissimilar
    except Exception as e:
        print(f"Error during similarity check: {e}")
        return 0  # Treat errors as dissimilar


def summarize_reviews(reviews, sentiment_type, api_key):
    """Summarizes positive or negative reviews using Generative AI."""
    if not reviews:
        return "No reviews to summarize.", "N/A"

    prompt = f"""You are an expert in summarizing restaurant reviews. Create a concise summary of the following {sentiment_type} restaurant reviews, highlighting key themes and recurring issues. Format the summary as a list of points, each focusing on a specific aspect like food quality, service, ambiance, etc.  Also, provide a justification for each point in your summary, citing specific excerpts or quotes from the reviews to support your claim.

Here are the {sentiment_type} reviews:
{reviews}"""

    try:
        genai.configure(api_key=api_key)  # Configure Gemini API *within* the function
        model = genai.GenerativeModel('gemini-2.0-flash')
        response = model.generate_content(prompt)
        return response.text.strip(), prompt
    except Exception as e:
        print(f"Error during API call for summarization: {e}")
        return "Error generating summary.", "API Error"


def analyze_competition(my_reviews, competitor_reviews, my_name, competitor_name, api_key):
    """Analyzes reviews to identify areas where each restaurant excels."""

    if not my_reviews or not competitor_reviews:
        return "Insufficient data for comparison.", "N/A", "Insufficient data for comparison.", "N/A"

    prompt_better = f"""You are an expert restaurant reviewer. Compare the following customer reviews of {my_name} with customer reviews of {competitor_name}. Identify specific aspects where {my_name} excels compared to {competitor_name}, such as food quality, service, ambiance, or value. Provide specific examples from the reviews as justification.

{my_name} Reviews:\n{my_reviews}\n\n{competitor_name} Reviews:\n{competitor_reviews}

Focus on highlighting the strengths of {my_name} based on these reviews."""

    prompt_worse = f"""You are an expert restaurant reviewer. Compare the following customer reviews of {my_name} with customer reviews of {competitor_name}. Identify specific aspects where {competitor_name} excels compared to {my_name}, such as food quality, service, ambiance, or value. Provide specific examples from the reviews as justification.

{my_name} Reviews:\n{my_reviews}\n\n{competitor_name} Reviews:\n{competitor_reviews}

Focus on highlighting the strengths of {competitor_name} based on these reviews."""

    try:
        genai.configure(api_key=api_key)  # Configure Gemini API *within* the function
        model = genai.GenerativeModel('gemini-2.0-flash')

        response_better = model.generate_content(prompt_better)
        my_better = response_better.text.strip()

        response_worse = model.generate_content(prompt_worse)
        competitor_better = response_worse.text.strip()

        return my_better, prompt_better, competitor_better, prompt_worse

    except Exception as e:
        print(f"Error during competition analysis: {e}")
        return "Error analyzing competition.", "API Error", "Error analyzing competition.", "API Error"


def analyze_trend_shift(previous_month_reviews, current_month_reviews, outlet, trend_type, api_key):
    """Analyzes reviews to identify reasons for trend shifts."""

    if not previous_month_reviews or not current_month_reviews:
        return "Insufficient data to determine trend shifts."

    prompt = f"""You are an expert in analyzing restaurant review trends. Identify the top 3 reasons why customer reviews for {outlet} shifted from {trend_type.split('_')[1]} in the previous month to {trend_type.split('_')[3]} in the current month. Provide specific examples from the reviews as justification.

Previous Month Reviews:\n{previous_month_reviews}\n\nCurrent Month Reviews:\n{current_month_reviews}

Focus on the most significant changes in customer sentiment."""

    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-2.0-flash')
        response = model.generate_content(prompt)
        return response.text.strip()

    except Exception as e:
        print(f"Error during trend shift analysis: {e}")
        return "Error analyzing trend shift."


def get_month_reviews(cursor, outlet, month, review_text_col_name, review_sentiment_col_name):
    """Retrieves positive and negative reviews for a specific outlet and month."""
    month_reviews = defaultdict(list)
    query = f"""
        SELECT {review_text_col_name}, {review_sentiment_col_name}
        FROM reviews_trend_dummy
        WHERE outlet = %s AND review_month = %s
    """
    cursor.execute(query, (outlet, month))
    rows = cursor.fetchall()
    for review_text, review_sentiment in rows:
        if isinstance(review_sentiment, str):
            sentiment_lower = review_sentiment.lower()
            if sentiment_lower == 'positive':
                month_reviews['positive'].append(review_text)
            elif sentiment_lower == 'negative':
                month_reviews['negative'].append(review_text)
    return month_reviews

def process_reviews_and_store_data(api_key):
    """Processes reviews from the database and stores aggregated data in output_dummy_2."""

    competitors = {
        # 'South Plainfield': 'Chand Palace', #Competitors to Add or Remove.
        # 'Princeton': 'Saravana Bhavan', #Competitors to Add or Remove.
        'Parsippany': 'Sangeetha', #Competitors to Add or Remove.
        # 'Chicago': 'Udupi Palace' #Competitors to Add or Remove.
    }

    try:
        cnx = mysql.connector.connect(**db_config)
        cursor = cnx.cursor()

        # ***FIX: Getting column indices for main sheet and outlet***
        # Assuming the input table has columns 'outlet', 'reviews', 'review_sentiment',
        # 'dish_sentiment', 'staff_sentiment', 'category_sentiment'
        outlet_col_name = 'outlet'
        review_text_col_name = 'reviews'
        review_sentiment_col_name = 'review_sentiment'  # Make sure it's the correct name
        dish_sentiment_col_name = 'dish_sentiment'
        staff_sentiment_col_name = 'staff_sentiment'
        category_sentiment_col_name = 'category_sentiment'
        month_col_name = 'review_month' #Change here

        for outlet, competitor in competitors.items(): #Update The For loop to be Key Value Pair
            print(f"Processing outlet: {outlet} and competitor: {competitor}")

            #Get all reviews of the outlet, we need to create our buckets
            allReviews = defaultdict(lambda: {'positive': [], 'negative': []})

            dish_positive_counts = defaultdict(int)
            dish_negative_counts = defaultdict(int)
            staff_positive_counts = defaultdict(int)
            staff_negative_counts = defaultdict(int)
            category_positive_counts = defaultdict(int)
            category_negative_counts = defaultdict(int)
            
            #Get data for the competitor.
            allCompetitorReviews = defaultdict(lambda: {'positive': [], 'negative': []})
            
            #Select all data for month 1 and 2 for both the competitor and my outlet.

            select_query = f"""
                SELECT {outlet_col_name}, {review_text_col_name}, {review_sentiment_col_name},
                       {dish_sentiment_col_name}, {staff_sentiment_col_name}, {category_sentiment_col_name}, {month_col_name}
                FROM reviews_trend_dummy
                WHERE ({outlet_col_name} = %s OR {outlet_col_name} = %s) AND {month_col_name} IN (1, 2) #Restricting it to months 1 and 2.
            """

            cursor.execute(select_query, (outlet, competitor)) #Passing in the outlet followed by competitor.
            review_rows = cursor.fetchall()

            for row in review_rows:
                try:
                    current_outlet, review_text, review_sentiment, dish_sentiment_json, staff_sentiment_json, category_sentiment_json, month = row
                    
                    if isinstance(review_sentiment, str):
                        sentiment_lower = review_sentiment.lower()

                        #Handle Reviews and their sentiments
                        if current_outlet == outlet:
                            if sentiment_lower == 'positive':
                                allReviews[current_outlet]['positive'].append(review_text)

                            elif sentiment_lower == 'negative':
                                allReviews[current_outlet]['negative'].append(review_text)

                        elif current_outlet == competitor: #Now we need to load the reviews for the competitor.
                            if sentiment_lower == 'positive':
                                allCompetitorReviews[current_outlet]['positive'].append(review_text)
                            elif sentiment_lower == 'negative':
                                allCompetitorReviews[current_outlet]['negative'].append(review_text)
                    
                    #Load the JSON for the next variable
                    try:
                        dish_sentiment = json.loads(dish_sentiment_json) if dish_sentiment_json else {}
                        staff_sentiment = json.loads(staff_sentiment_json) if staff_sentiment_json else {}
                        category_sentiment = json.loads(category_sentiment_json) if category_sentiment_json else {}

                    except (json.JSONDecodeError, TypeError) as e:
                        print(f"Failed to decode JSON: {e}")
                        dish_sentiment = {}
                        staff_sentiment = {}
                        category_sentiment = {}

                    # Update counts based on sentiments.
                    #We are only adding for our own outlet not the competitor.
                    if current_outlet == outlet:
                        for dish, sentiment in dish_sentiment.items():
                            if sentiment == 'positive':
                                dish_positive_counts[dish] += 1
                            elif sentiment == 'negative':
                                dish_negative_counts[dish] += 1

                        for staff, sentiment in staff_sentiment.items():
                            if sentiment == 'positive':
                                staff_positive_counts[staff] += 1
                            elif sentiment == 'negative':
                                staff_negative_counts[staff] += 1

                        for category, sentiment in category_sentiment.items():
                            if sentiment == 'positive':
                                category_positive_counts[category] += 1
                            elif sentiment == 'negative':
                                category_negative_counts[category] += 1

                except Exception as e:
                    print(f"Error collecting review data in row: {e}")
                    traceback.print_exc()
                    continue
            
            #Checking if there are reviews for analysis.

            if not allReviews or not allReviews[outlet]['positive'] or not allReviews[outlet]['negative']:
                print(f"Skipping outlet {outlet}: No reviews found.")
                continue #Skip if there is not previous review

            #We have data, so Lets create variables
            overall_positive_count = len(allReviews[outlet]['positive'])
            overall_negative_count = len(allReviews[outlet]['negative'])
            overall_neutral_count = 0  # Assuming unused, setting to 0
            
            dish_positive_counts_aggregated = aggregate_counts(dish_positive_counts, api_key)
            dish_negative_counts_aggregated = aggregate_counts(dish_negative_counts, api_key)
            staff_positive_counts_aggregated = aggregate_counts(staff_positive_counts, api_key)
            staff_negative_counts_aggregated = aggregate_counts(staff_negative_counts, api_key)
            category_positive_counts_aggregated = aggregate_counts(category_positive_counts, api_key)
            category_negative_counts_aggregated = aggregate_counts(category_negative_counts, api_key)
            
            positive_reviews_for_summary = allReviews[outlet]['positive']
            negative_reviews_for_summary = allReviews[outlet]['negative']
            
            # Generate summaries
            positive_summary, pos_summary_justification = summarize_reviews("\n".join(positive_reviews_for_summary), "positive", api_key)
            negative_summary, neg_summary_justification = summarize_reviews("\n".join(negative_reviews_for_summary), "negative", api_key)

            # Analyze competition (skipping for now)
            my_better, my_better_justification, competitor_better, competitor_better_justification = analyze_competition(allReviews[outlet]['positive'] + allReviews[outlet]['negative'], allCompetitorReviews[competitor]['positive'] + allCompetitorReviews[competitor]['negative'], outlet, competitor, api_key)

            #Get trends.
            trend_pos_to_neg = ""
            trend_neg_to_pos = ""
            
            # Load month reviews to make sure the logic works
            month1_reviews = get_month_reviews(cursor, outlet, 1, review_text_col_name, review_sentiment_col_name)
            month2_reviews = get_month_reviews(cursor, outlet, 2, review_text_col_name, review_sentiment_col_name)

            #We're just doing trend from pos->neg for brevity
            trend_pos_to_neg = analyze_trend_shift(month1_reviews['positive'], month2_reviews['negative'], outlet, "Positive_to_Negative", api_key) #call the trend analysis for a positive to negative analysis with month 1 and the over reviews of month 2.
            
            # Prepare the SQL query for insertion
            # Update the insert queries, since we are now not using insert
            update_query = """
                INSERT INTO output_dummy_2 (
                    outlet, review_month, overall_positive_count, overall_negative_count, overall_neutral_count,
                    dish_positive_counts, dish_negative_counts, staff_positive_counts, staff_negative_counts,
                    category_positive_counts, category_negative_counts, positive_summary,
                    pos_summary_justification, negative_summary, neg_summary_justification,
                    where_i_do_better, where_i_do_better_justification,
                    where_competitor_do_better, where_competitor_do_better_justification,
                    trend_pos_to_neg, trend_neg_to_pos
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """ 

            # Prepare data for insertion
            data_to_insert = (
                outlet, 2, overall_positive_count, overall_negative_count, overall_neutral_count,
                json.dumps(dish_positive_counts_aggregated), json.dumps(dish_negative_counts_aggregated),
                json.dumps(staff_positive_counts_aggregated), json.dumps(staff_negative_counts_aggregated),
                json.dumps(category_positive_counts_aggregated), json.dumps(category_negative_counts_aggregated),
                positive_summary, pos_summary_justification, negative_summary, neg_summary_justification,
                my_better, my_better_justification, competitor_better, competitor_better_justification,
                trend_pos_to_neg, trend_neg_to_pos
            )

            # Execute the SQL query
            cursor.execute(update_query, data_to_insert)
            cnx.commit()  # Commit the changes to the database.

            print(f"Data for outlet {outlet} inserted successfully.")

        print("All reviews processed and inserted into output_dummy_2.")

    except mysql.connector.Error as err:
        print(f"Error connecting to MySQL: {err}")
        traceback.print_exc()

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        traceback.print_exc()

    finally:
        if cursor:
            cursor.close()
        if cnx:
            cnx.close()