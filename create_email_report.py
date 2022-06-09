import argparse
import csv
import json
import os
import requests
import sys
import traceback
from dotenv import load_dotenv

VERBOSE = False


def get_api_key() -> None:
    """Retrieve the NGP API key and store it globally."""
    load_dotenv()
    global API_KEY
    API_KEY = os.getenv('MIKES_API_KEY')


def call_api(path) -> dict:
    """Make a request to the given path and return the content as json."""
    base_url = 'https://api.myngp.com/v2/broadcastEmails/'
    url = base_url + path
    headers = {'Accept': 'application/json'}
    response = requests.get(url, headers=headers, auth=('apiuser', API_KEY))
    return json.loads(response.text)


def get_email_ids() -> list:
    """Obtain all email ids from the server."""
    emails = call_api('')['items']
    return [email['emailMessageId'] for email in emails]


def get_email_stats(email_id: int) -> tuple:
    """Obtain an email's statistics and variants from the server and order them."""
    email_info = call_api(str(email_id) + '?$expand=statistics')
    email_stats = email_info['statistics']
    ordered_stats = [email_id, email_info['name']]
    for stat in ['recipients', 'opens', 'clicks', 'unsubscribes', 'bounces']:
        ordered_stats.append(email_stats[stat])
    return (email_info['variants'], ordered_stats)


def get_variant_stats(email_id: int, variant_id: int) -> dict:
    """Obtain a variant's statistics from the server."""
    path = str(email_id) + '/variants/' + str(variant_id) + '?$expand=statistics'
    variant_stats = call_api(path)['statistics']
    return variant_stats


def get_top_variant(email_id: int, variants: dict) -> str:
    """Iterate through all variants of 1 email to identify the top open ratio."""
    top_variant = ''
    top_ratio = 0.0
    top_recipients = 0
    if VERBOSE:
        print("-"*60)
        print('Email ID: ', email_id)
    for variant in variants:
        variant_stats = get_variant_stats(email_id, variant['emailMessageVariantId'])
        variant_recipients = variant_stats['recipients']
        variant_opens = variant_stats['opens']
        variant_ratio = variant_opens / variant_recipients
        if VERBOSE:
            print('TOP', top_variant, '||', top_ratio, '||', top_recipients)
            print('VAR', variant['name'], '||', variant_ratio, '||', variant_recipients)
        if variant_ratio > top_ratio or \
                (variant_ratio == top_ratio and variant_recipients > top_recipients):
            top_variant = variant['name']
            top_ratio = variant_ratio
            top_recipients = variant_recipients
    if VERBOSE:
        print('WIN', top_variant, '||', top_ratio, '||', top_recipients)
    return top_variant


def main(be_verbose: bool) -> int:
    try:
        # Set verbosity
        if be_verbose:
            global VERBOSE
            VERBOSE = True
        # Load the API key
        get_api_key()
        # Get and sort the email ids
        email_ids = sorted(get_email_ids())
        # Construct matrix to hold report data
        email_report = [[] for email_id in email_ids]
        # For each email, add statistics and top variant name to matrix
        for index, email_id in enumerate(email_ids):
            variants, email_stats = get_email_stats(email_id)
            top_variant = get_top_variant(email_id, variants)
            email_stats.append(top_variant)
            email_report[index] = email_stats
        # Write header and report data to file
        with open('EmailReport.csv', 'w', newline='') as csvfile:
            csvwriter = csv.writer(csvfile)
            csvwriter.writerow([
                'Message ID', 'Email Name', 'Recipients', 'Opens', 'Clicks', 
                'Unsubscribes', 'Bounces', 'Top Variant'])
            csvwriter.writerows(email_report)
        print('Email report complete; file is EmailReport.csv')
        return 0
    except Exception as e:
        print('Exception: ')
        print("-"*60)
        traceback.print_exc(file=sys.stdout)
        print("-"*60)
        return 1


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="This script will generate a report of Emails and their top Variants.")
    parser.add_argument('-v', '--verbose', action='store_true', default=False)
    args = parser.parse_args()
    sys.exit(main(args.verbose)) 
