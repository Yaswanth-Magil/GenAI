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
# db_config = {
#     'user': 'root',
#     'password': 'Yaswanth123.',
#     'host': 'localhost',
#     'database': 'genai',
#     'raise_on_warnings': True
# }

db_config = {
    'user': 'root',
    'password': 'Z*ZlRmnFCP@9V',
    'host': '10.162.0.3',
    'database': 'mhrq',
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
    """Summarizes positive or negative reviews using Generative AI with specific prompts."""
    if not reviews:
        return "No reviews to summarize.", "N/A"

    if sentiment_type == "positive":
        prompt = f"""Summarize the positive reviews concisely in 5 to 10 points.
Each point should cover a separate topic or key concept.
Provide one observation per point, ensuring it is limited to one idea.
Do not include any additional information or comments.
Format:
[Key Concept]: [Observation or feature]. and Task: Summarize positive Restaurant Reviews

Example:
Delicious and Authentic South Indian Food: The food quality consistently receives high praise, with reviewers highlighting the authentic taste and deliciousness of various South Indian dishes. 
Wide Variety of Menu Options, Including Northern and Southern Indian Dishes: The restaurant offers a diverse menu, catering to different tastes and preferences, spanning both North and South Indian cuisines. 
Exceptional Service, Particularly by Specific Staff Members: Several reviewers specifically mention the outstanding service provided by staff members like Niyas, Abdul, Farzi, and Shasny, praising their attentiveness, recommendations, and accommodating nature. 
Clean and Well-Maintained Atmosphere: The cleanliness and ambiance of the restaurant contribute to a pleasant dining experience. 
Consistently Positive Experience and High Likelihood of Return: Reviewers express a strong desire to revisit the restaurant, indicating a high level of satisfaction with their overall experience. 
Overall Positive Vegetarian Dining Experience: The restaurant is regarded as a top choice for vegetarian cuisine, offering a satisfying experience for those seeking quality vegetarian options.

Here are the positive reviews:
{reviews}"""
    elif sentiment_type == "negative":
        prompt = f"""Summarize the negative reviews concisely in 5 to 10 points.
Each point should cover a separate topic or key concept.
Provide one observation per point, ensuring it is limited to one idea.
Include a justification or example review for each point.
Do not include any additional information or comments.
Format:
[Key Concept]: [Observation or feature]. and Task: Summarize negative Restaurant Reviews

Example:
Inconsistent Food Quality and Taste: The food quality appears to be inconsistent and often disappointing. A longtime patron also stated, "This restaurant used to be our longtime favorite, but our last couple of visits have been disappointing." 
Overpricing and Poor Value: The prices are perceived as too high for the quality and portion sizes offered. Example: "Horrible food... The food is overpriced and lacking in quality. For these prices, you should get way better." 
Poor and Inconsistent Service: Service issues are a recurring complaint, including rudeness, slow service, and unequal attention to diners at the same table. One reviewer also noted being made to wait longer than expected: "Made us wait for 40 mins after calling us on 20 mins." 
Crowded and Uncomfortable Ambiance: The restaurant is often crowded with tight seating arrangements, contributing to an unpleasant dining experience. Literally every table was empty." imply a problem with space and potentially seating arrangements. 
Inconsistency with Spice Levels: The spice levels in dishes are inconsistent and can be unexpectedly high, causing discomfort. "Most recently, on January 26th, we were served a masala dosa with an extremely spicy potato filling, far spicier than usual." 
Overall Disappointment and Negative Recommendation: Many reviewers express overall disappointment and explicitly state they will not return. "Highly overrated and a disappointing experience!" and "Never ever going back." and "Better go to SARAVANA BHAVAN which is million times BETTER." indicate a strong negative sentiment and recommendation against visiting the restaurant.'.

Here are the negative reviews:
{reviews}"""
    else:
        return "Invalid sentiment type.", "N/A"

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

    prompt_better = f"""You are an expert restaurant reviewer. Compare the following customer reviews of {my_name} with customer reviews of {competitor_name}.
Identify specific aspects where {my_name} excels compared to {competitor_name}, such as food quality, service, ambiance, or value.
Do not include any additional information or comments. Do not provide additional information otherthan asked.
While mentioning my name, use like this `A2B {my_name}` . Stick on to the format provided. Do not provide anything excess other than asked.
Use the following format:
Format: (Make the key concept bold with **)
Key Concept: [Observation or feature].
Example:
Breadth of Menu: Chand Palace typically boasts a larger menu encompassing a wider variety of North Indian, South Indian, and Indo-Chinese dishes.
Consistent Quality and Availability: Being a more established chain, Chand Palace generally offers more consistent quality and availability across its various locations.
Accessibility and Convenience: The chain establishment of Chand Palace provides more locations for customers to choose from for easier access to its menu.
"""

    prompt_worse = f"""You are an expert restaurant reviewer. Compare the following customer reviews of {my_name} with customer reviews of {competitor_name}.
Identify specific aspects where {competitor_name} excels compared to {my_name}, such as food quality, service, ambiance, or value.
Do not include any additional information or comments. Do not provide additional information otherthan asked.
While mentioning my name, use like this `A2B {my_name}` . Stick on to the format provided. Do not provide anything excess other than asked.
Use the following format:
Format: (Make the key concept bold with **)
Key Concept: [Observation or feature].
Example:
Price: Depending on the specific restaurant, some South Plainfield establishments might offer slightly lower prices on certain items compared to Chand Palace.
Specific Regional Dishes: If you are looking for hyper-local or specific regional Indian dishes, smaller restaurants in South Plainfield might specialize in those areas, while Chand Palace caters to a broader, pan-Indian palate.
Takeout/Delivery Speed: Depending on proximity and staffing, certain South Plainfield restaurants *might* offer slightly faster takeout/delivery options for residents in that specific area.
"""
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


def analyze_trend_shift(previous_month_reviews, current_month_reviews, outlet, previous_sentiment, current_sentiment, api_key):
    """Analyzes reviews to identify reasons for trend shifts."""

    if not previous_month_reviews or not current_month_reviews:
        return "Insufficient data to determine trend shifts."

    prompt = f"""Review Changes from Positive to Negative: Note any reviews that shifted from positive to negative in subsequent months.


Use the following format:
Format: (Make the key concept bold with **)
Key Concept: [Observation or feature].

Example: 
Food Consistency Issues: 
Previously, the food was praised for being consistently good across locations, maintaining authentic South Indian flavors. 
This month, reviews highlight inconsistent food quality, with some customers stating that they had great experiences before, but recent visits were disappointing. 

Spice Level Issues: 
Last month, the restaurant was praised for offering customizable spice levels and catering to dietary preferences. 
This month, there are complaints that spice levels are inconsistent, sometimes excessively high, making dishes difficult to enjoy. 

Ambiance Perception Declined: 
Previously, the ambiance was considered pleasant, with instrumental music adding to the experience. 
Now, there are complaints that the restaurant is overcrowded, seating is tight, and the atmosphere is uncomfortable. 
Service Deterioration: 
While last month’s reviews consistently praised the friendly and efficient service, this month, multiple complaints mention rudeness, long wait times, and unprofessional behavior. 


Identify the top 3 reasons why customer reviews for {outlet} shifted from {previous_sentiment} in the previous month to {current_sentiment} in the current month. Provide specific examples from the reviews as justification.

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
            elif sentiment_lower != 'positive' and sentiment_lower != 'negative':
                month_reviews['neutral'].append(review_text)
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
"Note: Positive reviews consistently form the majority across all months. Negative feedback peaks in November at 27.1% before decreasing in December at 17.5% and increasing in January to 11.8%, while the overall sentiment depicted in the pie chart shows 86.3% positive, 11.8% negative, and 2.0% neutral."

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
"Note: Overall Experience and Food Quality received the highest positive mentions (33). However, Overall Experience also had the most negative feedback (12). Value for Money received the lowest positive mentions (3), this category is a point of concern. Staff Friendliness and Service also scored well with positive feedback. All categories, except for Value for Money, received more positive than negative feedback."

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
        'South Plainfield': 'Chand Palace',  # Competitors to Add or Remove.
        # 'Princeton': 'Saravana Bhavan',  # Competitors to Add or Remove.
        # 'Parsippany': 'Sangeetha',      # Competitors to Add or Remove.
        # 'Chicago': 'Udupi Palace'      # Competitors to Add or Remove.
    }

    try:
        cnx = mysql.connector.connect(**db_config)
        cursor = cnx.cursor()

        # ***FIX: Getting column indices for main sheet and outlet***
        outlet_col_name = 'Outlet'
        source_col_name = 'Source'
        date_col_name = 'Date'
        review_month_col_name = 'review_month'
        year_col_name = 'Year'
        review_text_col_name = 'Reviews'
        review_sentiment_col_name = 'review_sentiment'
        dish_sentiment_col_name = 'dish_sentiment'
        staff_sentiment_col_name = 'staff_sentiment'
        category_sentiment_col_name = 'category_sentiment'


        for outlet, competitor in competitors.items():
            print(f"Processing outlet: {outlet} and competitor: {competitor}")

            # Get all reviews for the outlet for the last 3 months
            # First, get all months available
            cursor.execute(f"SELECT DISTINCT `{review_month_col_name}` FROM reviews_trend_dummy WHERE `{outlet_col_name}` = %s ORDER BY `{review_month_col_name}` DESC", (outlet,))
            available_months = [row[0] for row in cursor.fetchall()]

            # Filter available_months to only include months less than or equal to the adjustable month
            available_months = [month for month in available_months if month <= month_to_process]
            # Get the last 3 months, making sure to include the manually adjustable month
            months_to_process = sorted(available_months[:3])  # Ensure we process in ascending order

            print(f"Months to process: {months_to_process}")

            # Prepare to store review data for trend analysis
            review_data_for_trend = []

            print(f"Processing month: {month_to_process}")  # Added printing of the Processing Month

            # Check if the record already exists BEFORE potentially skipping
            check_query = "SELECT * FROM output_dummy_2 WHERE outlet = %s AND review_month = %s"
            cursor.execute(check_query, (outlet, month_to_process))
            existing_record = cursor.fetchone()

            allReviews = defaultdict(lambda: {'positive': [], 'negative': [], 'neutral': []})
            dish_positive_counts = defaultdict(int)
            dish_negative_counts = defaultdict(int)
            staff_positive_counts = defaultdict(int)
            staff_negative_counts = defaultdict(int)
            category_positive_counts = defaultdict(int)
            category_negative_counts = defaultdict(int)
            allCompetitorReviews = defaultdict(lambda: {'positive': [], 'negative': []})

            select_query = f"""
                SELECT `{outlet_col_name}`,`{review_text_col_name}`, `{review_sentiment_col_name}`,
                       `{dish_sentiment_col_name}`, `{staff_sentiment_col_name}`, `{category_sentiment_col_name}`, `{review_month_col_name}`, `{year_col_name}`
                FROM reviews_trend_dummy
                WHERE (`{outlet_col_name}` = %s OR `{outlet_col_name}` = %s) AND `{review_month_col_name}` = %s
            """

            cursor.execute(select_query, (outlet, competitor, month_to_process))
            review_rows = cursor.fetchall()

            for row in review_rows:
                try:
                    Outlet,Reviews,review_sentiment,dish_sentiment,staff_sentiment,category_sentiment,review_month,Year  = row
                    if isinstance(review_sentiment, str):
                        sentiment_lower = review_sentiment.lower()

                        if Outlet == outlet:
                            if sentiment_lower == 'positive':
                                allReviews[Outlet]['positive'].append(Reviews)

                            elif sentiment_lower == 'negative':
                                allReviews[Outlet]['negative'].append(Reviews)
                            elif sentiment_lower != 'positive' and sentiment_lower != 'negative':
                                allReviews[Outlet]['neutral'].append(Reviews)


                        elif Outlet == competitor:
                            if sentiment_lower == 'positive':
                                allCompetitorReviews[Outlet]['positive'].append(Reviews)
                            elif sentiment_lower == 'negative':
                                allCompetitorReviews[Outlet]['negative'].append(Reviews)

                        try:
                            dish_sentiment = json.loads(str(dish_sentiment)) if dish_sentiment else {}
                            staff_sentiment = json.loads(str(staff_sentiment)) if staff_sentiment else {}
                            category_sentiment = json.loads(str(category_sentiment)) if category_sentiment else {}

                        except (json.JSONDecodeError, TypeError) as e:
                            print(f"Failed to decode JSON: {e}")
                            dish_sentiment = {}
                            staff_sentiment = {}
                            category_sentiment = {}

                        if Outlet == outlet:
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
            overall_neutral_count = len(allReviews[outlet]['neutral'])  # Change made here to add neutral count.

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

            # Trend Shift Analysis
            trend_pos_to_neg = ""
            trend_neg_to_pos = ""

            # Added trend shift logic here
            if len(months_to_process) > 1 and month_to_process == months_to_process[-1]:  # Ensure we have at least two months and it's the latest
                previous_month = months_to_process[-2]
                print(f"Analyzing trend shift from {previous_month} to {month_to_process}")

                # Retrieve the previous month's reviews
                previous_month_reviews = get_month_reviews(cursor, outlet, previous_month, review_text_col_name, review_sentiment_col_name)
                current_month_reviews = get_month_reviews(cursor, outlet, month_to_process, review_text_col_name, review_sentiment_col_name)

                # Analyze shift from positive to negative
                if previous_month_reviews['positive'] and current_month_reviews['negative']:
                    trend_pos_to_neg = analyze_trend_shift(
                        "\n".join(previous_month_reviews['positive']),
                        "\n".join(current_month_reviews['negative']),
                        outlet, "positive", "negative", api_key
                    )
                    print(f"Trend from Positive to Negative: {trend_pos_to_neg}")

                # Analyze shift from negative to positive
                if previous_month_reviews['negative'] and current_month_reviews['positive']:
                    trend_neg_to_pos = analyze_trend_shift(
                        "\n".join(previous_month_reviews['negative']),
                        "\n".join(current_month_reviews['positive']),
                        outlet, "negative", "positive", api_key
                    )
                    print(f"Trend from Negative to Positive: {trend_neg_to_pos}")

            # Generate the category note
            category_note = generate_category_note(outlet, category_positive_counts_aggregated, category_negative_counts_aggregated, api_key)

            # Prepare the data for trend analysis for current month
            review_data_for_trend.append({
                'review_month': month_to_process,
                'overall_positive_count': overall_positive_count,
                'overall_negative_count': overall_negative_count,
                'overall_neutral_count': overall_neutral_count
            })

            # Generate the trend note only if the month is the manually adjustable month.
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
                    trend_pos_to_neg, trend_neg_to_pos, trend_note, category_note,
                    outlet, review_month
                )
                try:
                    cursor.execute(update_query, data_to_update)
                    cnx.commit()
                    print(f"Data for outlet {outlet} and month {month_to_process} updated successfully.")
                except Exception as e:
                    print(f"An error occurred in the update statement: {e}")
                    traceback.print_exc()
            else:
                # Prepare the SQL query for insertion
                insert_query = """
                    INSERT INTO output_dummy_2 (
                        outlet, review_month, year, overall_positive_count, overall_negative_count, overall_neutral_count,
                        dish_positive_counts, dish_negative_counts, staff_positive_counts, staff_negative_counts,
                        category_positive_counts, category_negative_counts, positive_summary,
                        pos_summary_justification, negative_summary, neg_summary_justification,
                        where_i_do_better, where_i_do_better_justification,
                        where_competitor_do_better, where_competitor_do_better_justification,
                        trend_pos_to_neg, trend_neg_to_pos, trend_note, category_note
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """
                data_to_insert = (
                    Outlet, review_month, int(Year), overall_positive_count, overall_negative_count, overall_neutral_count,
                    json.dumps(dish_positive_counts_aggregated), json.dumps(dish_negative_counts_aggregated),
                    json.dumps(staff_positive_counts_aggregated), json.dumps(staff_negative_counts_aggregated),
                    json.dumps(category_positive_counts_aggregated), json.dumps(category_negative_counts_aggregated),
                    positive_summary, pos_summary_justification, negative_summary, neg_summary_justification,
                    my_better, my_better_justification, competitor_better, competitor_better_justification,
                    trend_pos_to_neg, trend_neg_to_pos, trend_note, category_note
                )
                print(data_to_insert)

                # Execute the SQL query
                try:
                    cursor.execute(insert_query, data_to_insert)
                    cnx.commit()
                    print(f"Data for outlet {outlet} and month {month_to_process} inserted successfully.")
                except Exception as e:
                    print(f"An error occurred in the insert statement: {e}")
                    traceback.print_exc()

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