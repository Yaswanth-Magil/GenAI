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

    # Modify trend_type handling to prevent IndexError
    trend_parts = trend_type.split('_')
    if len(trend_parts) < 4:
        print(f"Warning: Invalid trend_type format: {trend_type}. Skipping trend analysis.")
        return "Invalid trend_type format."

    previous_sentiment = trend_parts[1]
    current_sentiment = trend_parts[3]

    prompt = f"""You are an expert in analyzing restaurant review trends. Identify the top 3 reasons why customer reviews for {outlet} shifted from {previous_sentiment} in the previous month to {current_sentiment} in the current month. Provide specific examples from the reviews as justification.

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

def generate_trend_note(outlet, review_data, api_key):
    """Generates a trend note for the visualization based on the last 3 months of review data."""

    # Sort review data by review_month in descending order (latest months first)
    sorted_review_data = sorted(review_data, key=lambda x: x['review_month'], reverse=True)

    # Limit to the latest 3 months or fewer
    latest_months_data = sorted_review_data[:min(3, len(sorted_review_data))]

    if not latest_months_data:
        return "Insufficient data to generate a trend note."

    # Prepare data for the prompt
    month_notes = []
    for month_data in latest_months_data:
        month = month_data['review_month']
        overall_positive = month_data['overall_positive_count']
        overall_negative = month_data['overall_negative_count']
        overall_neutral = month_data['overall_neutral_count']
        total_reviews = overall_positive + overall_negative + overall_neutral

        # Calculate percentages
        positive_percentage = (overall_positive / total_reviews) * 100 if total_reviews > 0 else 0
        negative_percentage = (overall_negative / total_reviews) * 100 if total_reviews > 0 else 0
        neutral_percentage = (overall_neutral / total_reviews) * 100 if total_reviews > 0 else 0

        month_notes.append({
            'month': month,
            'positive_percentage': positive_percentage,
            'negative_percentage': negative_percentage,
            'neutral_percentage': neutral_percentage
        })

    # Construct prompt
    prompt = f"""You are an expert in creating concise notes for restaurant review data visualizations. Given the following restaurant review data for {outlet} over the last {len(latest_months_data)} months, create a single note to describe trends in overall sentiment (positive, negative, and neutral reviews).
Your note should be formatted like the example below, comparing the percentages of positive, negative, and neutral reviews across the months. Identify any significant peaks or dips in sentiment.

Example Note:
"Positive reviews consistently form the majority across all months. Negative feedback peaks in November at 27.1% before decreasing in December at 17.5% and increasing in January to 11.8%, while the overall sentiment depicted in the pie chart shows 86.3% positive, 11.8% negative, and 2.0% neutral."

Here is the review data:
"""
    for month_note in month_notes:
        prompt += f"""Month: {month_note['month']}, Positive: {month_note['positive_percentage']:.1f}%, Negative: {month_note['negative_percentage']:.1f}%, Neutral: {month_note['neutral_percentage']:.1f}%\n"""

    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-2.0-flash')
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        print(f"Error during trend note generation: {e}")
        return "Error generating trend note."


def generate_category_note(outlet, category_positive_counts_aggregated, category_negative_counts_aggregated, api_key):
    """Generates a category note for the visualization based on aggregated category sentiment data."""

    prompt = f"""You are an expert in creating concise notes for restaurant review data visualizations. Given the following aggregated sentiment data for different categories of customer feedback for {outlet}, create a single note to describe the distribution of positive and negative feedback across the categories.
Your note should be formatted like the example below, mentioning the categories with the highest positive and negative mentions, as well as any categories that may be a point of concern (low positive feedback).

Example Note:
"Overall Experience and Food Quality received the highest positive mentions (33). However, Overall Experience also had the most negative feedback (12). Value for Money received the lowest positive mentions (3), this category is a point of concern. Staff Friendliness and Service also scored well with positive feedback. All categories, except for Value for Money, received more positive than negative feedback."

Here is the category sentiment data:
Positive: {json.dumps(category_positive_counts_aggregated)}
Negative: {json.dumps(category_negative_counts_aggregated)}
"""

    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-2.0-flash')
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        print(f"Error during category note generation: {e}")
        return "Error generating category note."

def process_reviews_and_store_data(api_key, month_to_process=2):
    """Processes reviews from the database and stores aggregated data in output_dummy_2, including trend and category notes."""

    competitors = {
        'South Plainfield': 'Chand Palace', #Competitors to Add or Remove.
        # 'Princeton': 'Saravana Bhavan', #Competitors to Add or Remove.
        # 'Parsippany': 'Sangeetha', #Competitors to Add or Remove.
        # 'Chicago': 'Udupi Palace' #Competitors to Add or Remove.
    }

    try:
        cnx = mysql.connector.connect(**db_config)
        cursor = cnx.cursor()

        # ***FIX: Getting column indices for main sheet and outlet***
        outlet_col_name = 'outlet'
        review_text_col_name = 'reviews'
        review_sentiment_col_name = 'review_sentiment'
        dish_sentiment_col_name = 'dish_sentiment'
        staff_sentiment_col_name = 'staff_sentiment'
        category_sentiment_col_name = 'category_sentiment'
        month_col_name = 'review_month'

        for outlet, competitor in competitors.items():
            print(f"Processing outlet: {outlet} and competitor: {competitor}")

            # Get all reviews for the outlet for the last 3 months
            # First, get all months available
            cursor.execute(f"SELECT DISTINCT {month_col_name} FROM reviews_trend_dummy WHERE {outlet_col_name} = %s ORDER BY {month_col_name} DESC", (outlet,))
            available_months = [row[0] for row in cursor.fetchall()]

            # Filter available_months to only include months less than or equal to the adjustable month
            available_months = [month for month in available_months if month <= month_to_process]
            # Get the last 3 months, making sure to include the manually adjustable month
            months_to_process = sorted(available_months[:3])  # Ensure we process in ascending order

            # Prepare to store review data for trend analysis
            review_data_for_trend = []

            print(f"Processing month: {month_to_process}")  # Added printing of the Processing Month

            allReviews = defaultdict(lambda: {'positive': [], 'negative': []})
            dish_positive_counts = defaultdict(int)
            dish_negative_counts = defaultdict(int)
            staff_positive_counts = defaultdict(int)
            staff_negative_counts = defaultdict(int)
            category_positive_counts = defaultdict(int)
            category_negative_counts = defaultdict(int)
            allCompetitorReviews = defaultdict(lambda: {'positive': [], 'negative': []})

            select_query = f"""
                SELECT {outlet_col_name}, {review_text_col_name}, {review_sentiment_col_name},
                       {dish_sentiment_col_name}, {staff_sentiment_col_name}, {category_sentiment_col_name}, {month_col_name}
                FROM reviews_trend_dummy
                WHERE ({outlet_col_name} = %s OR {outlet_col_name} = %s) AND {month_col_name} = %s
            """

            cursor.execute(select_query, (outlet, competitor, month_to_process))
            review_rows = cursor.fetchall()

            for row in review_rows:
                try:
                    current_outlet, review_text, review_sentiment, dish_sentiment_json, staff_sentiment_json, category_sentiment_json, month = row

                    if isinstance(review_sentiment, str):
                        sentiment_lower = review_sentiment.lower()

                        if current_outlet == outlet:
                            if sentiment_lower == 'positive':
                                allReviews[current_outlet]['positive'].append(review_text)

                            elif sentiment_lower == 'negative':
                                allReviews[current_outlet]['negative'].append(review_text)

                        elif current_outlet == competitor:
                            if sentiment_lower == 'positive':
                                allCompetitorReviews[current_outlet]['positive'].append(review_text)
                            elif sentiment_lower == 'negative':
                                allCompetitorReviews[current_outlet]['negative'].append(review_text)

                        try:
                            dish_sentiment = json.loads(dish_sentiment_json) if dish_sentiment_json else {}
                            staff_sentiment = json.loads(staff_sentiment_json) if staff_sentiment_json else {}
                            category_sentiment = json.loads(category_sentiment_json) if category_sentiment_json else {}

                        except (json.JSONDecodeError, TypeError) as e:
                            print(f"Failed to decode JSON: {e}")
                            dish_sentiment = {}
                            staff_sentiment = {}
                            category_sentiment = {}

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

            if not allReviews or not allReviews[outlet]['positive'] or not allReviews[outlet]['negative']:
                print(f"Skipping outlet {outlet} for month {month_to_process}: No reviews found.")
                continue

            overall_positive_count = len(allReviews[outlet]['positive'])
            overall_negative_count = len(allReviews[outlet]['negative'])
            overall_neutral_count = 0

            dish_positive_counts_aggregated = aggregate_counts(dish_positive_counts, api_key)
            dish_negative_counts_aggregated = aggregate_counts(dish_negative_counts, api_key)
            staff_positive_counts_aggregated = aggregate_counts(staff_positive_counts, api_key)
            staff_negative_counts_aggregated = aggregate_counts(staff_negative_counts, api_key)
            category_positive_counts_aggregated = aggregate_counts(category_positive_counts, api_key)
            category_negative_counts_aggregated = aggregate_counts(category_negative_counts, api_key)

            positive_reviews_for_summary = allReviews[outlet]['positive']
            negative_reviews_for_summary = allReviews[outlet]['negative']

            positive_summary, pos_summary_justification = summarize_reviews("\n".join(positive_reviews_for_summary), "positive", api_key)
            negative_summary, neg_summary_justification = summarize_reviews("\n".join(negative_reviews_for_summary), "negative", api_key)

            my_better, my_better_justification, competitor_better, competitor_better_justification = analyze_competition(allReviews[outlet]['positive'] + allReviews[outlet]['negative'], allCompetitorReviews[competitor]['positive'] + allCompetitorReviews[competitor]['negative'], outlet, competitor, api_key)

            # Generate the category note
            category_note = generate_category_note(outlet, category_positive_counts_aggregated, category_negative_counts_aggregated, api_key)
            # Check if the record already exists
            check_query = "SELECT * FROM output_dummy_2 WHERE outlet = %s AND review_month = %s"
            cursor.execute(check_query, (outlet, month_to_process))
            existing_record = cursor.fetchone()

             # Prepare the data for trend analysis for current month

            review_data_for_trend.append({
                    'review_month': month_to_process,
                    'overall_positive_count': overall_positive_count,
                    'overall_negative_count': overall_negative_count,
                    'overall_neutral_count': overall_neutral_count
                })

            # Generate the trend note only if the month is the manually adjustable month. Only do this if the month matches
            trend_note = ""  # Initialize to an empty string
            if month_to_process == months_to_process[-1]:  # Check if it's the latest month
                 trend_note = generate_trend_note(outlet, review_data_for_trend, api_key)

            if existing_record:
                # Update the existing record
                update_query = """
                    UPDATE output_dummy_2
                    SET overall_positive_count = %s, overall_negative_count = %s, overall_neutral_count = %s,
                        dish_positive_counts = %s, dish_negative_counts = %s, staff_positive_counts = %s,
                        staff_negative_counts = %s, category_positive_counts = %s, category_negative_counts = %s,
                        positive_summary = %s, pos_summary_justification = %s, negative_summary = %s,
                        neg_summary_justification = %s, where_i_do_better = %s, where_i_do_better_justification = %s,
                        where_competitor_do_better = %s, where_competitor_do_better_justification = %s,
                        trend_pos_to_neg = %s, trend_neg_to_pos = %s, trend_note = %s, category_note = %s
                    WHERE outlet = %s AND review_month = %s
                """
                data_to_update = (
                    overall_positive_count, overall_negative_count, overall_neutral_count,
                    json.dumps(dish_positive_counts_aggregated), json.dumps(dish_negative_counts_aggregated),
                    json.dumps(staff_positive_counts_aggregated), json.dumps(staff_negative_counts_aggregated),
                    json.dumps(category_positive_counts_aggregated), json.dumps(category_negative_counts_aggregated),
                    positive_summary, pos_summary_justification, negative_summary, neg_summary_justification,
                    my_better, my_better_justification, competitor_better, competitor_better_justification,
                    "", "", trend_note, category_note,
                    outlet, month_to_process
                )
                cursor.execute(update_query, data_to_update)
                cnx.commit()
                print(f"Data for outlet {outlet} and month {month_to_process} updated successfully.")
            else:
                # Prepare the SQL query for insertion
                update_query = """
                    INSERT INTO output_dummy_2 (
                        outlet, review_month, overall_positive_count, overall_negative_count, overall_neutral_count,
                        dish_positive_counts, dish_negative_counts, staff_positive_counts, staff_negative_counts,
                        category_positive_counts, category_negative_counts, positive_summary,
                        pos_summary_justification, negative_summary, neg_summary_justification,
                        where_i_do_better, where_i_do_better_justification,
                        where_competitor_do_better, where_competitor_do_better_justification,
                        trend_pos_to_neg, trend_neg_to_pos, trend_note, category_note
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """

                # Prepare data for insertion
                data_to_insert = (
                    outlet, month_to_process, overall_positive_count, overall_negative_count, overall_neutral_count,
                    json.dumps(dish_positive_counts_aggregated), json.dumps(dish_negative_counts_aggregated),
                    json.dumps(staff_positive_counts_aggregated), json.dumps(staff_negative_counts_aggregated),
                    json.dumps(category_positive_counts_aggregated), json.dumps(category_negative_counts_aggregated),
                    positive_summary, pos_summary_justification, negative_summary, neg_summary_justification,
                    my_better, my_better_justification, competitor_better, competitor_better_justification,
                    "", "", trend_note, category_note
                )

                # Execute the SQL query
                cursor.execute(update_query, data_to_insert)
                cnx.commit()

                print(f"Data for outlet {outlet} and month {month_to_process} inserted successfully.")

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