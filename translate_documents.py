#!/usr/bin/env python3
"""
DeepL Document Translation Script
Translates month recap_EUR.docx into 6 languages using DeepL API
"""

import os
import sys
import io
import time
import requests
from pathlib import Path

# Fix encoding for Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Configuration
DEEPL_API_KEY = "374a8965-101a-4538-bc65-54506552650e"
SOURCE_FILE = r"C:\Users\Marco.Africani\Desktop\Month recap\month recap_EUR.docx"
OUTPUT_DIR = r"C:\Users\Marco.Africani\Desktop\Month recap\Translations"

# Target languages
LANGUAGES = {
    'DE': 'German',
    'ES': 'Spanish',
    'FR': 'French',
    'IT': 'Italian',
    'PT-PT': 'Portuguese',
    'RU': 'Russian'
}

# DeepL API endpoint (paid API, not free)
DEEPL_API_URL = "https://api.deepl.com/v2"


def upload_document(file_path, target_lang):
    """Upload document to DeepL for translation"""
    url = f"{DEEPL_API_URL}/document"

    with open(file_path, 'rb') as f:
        files = {'file': f}
        data = {
            'auth_key': DEEPL_API_KEY,
            'target_lang': target_lang,
            'source_lang': 'EN'  # Source is English
        }

        print(f"   Uploading document for {LANGUAGES[target_lang]} translation...")
        response = requests.post(url, files=files, data=data)

    if response.status_code == 200:
        result = response.json()
        return result['document_id'], result['document_key']
    else:
        raise Exception(f"Upload failed: {response.status_code} - {response.text}")


def check_translation_status(document_id, document_key):
    """Check if translation is complete"""
    url = f"{DEEPL_API_URL}/document/{document_id}"

    data = {
        'auth_key': DEEPL_API_KEY,
        'document_key': document_key
    }

    response = requests.post(url, data=data)

    if response.status_code == 200:
        result = response.json()
        return result['status']
    else:
        raise Exception(f"Status check failed: {response.status_code} - {response.text}")


def download_translated_document(document_id, document_key, output_path):
    """Download the translated document"""
    url = f"{DEEPL_API_URL}/document/{document_id}/result"

    data = {
        'auth_key': DEEPL_API_KEY,
        'document_key': document_key
    }

    print(f"   Downloading translated document...")
    response = requests.post(url, data=data)

    if response.status_code == 200:
        with open(output_path, 'wb') as f:
            f.write(response.content)
        return True
    else:
        raise Exception(f"Download failed: {response.status_code} - {response.text}")


def translate_document(source_file, target_lang, output_file):
    """Translate a document to target language"""
    print(f"\n{'='*80}")
    print(f"Translating to {LANGUAGES[target_lang]} ({target_lang})")
    print(f"{'='*80}")

    try:
        # Step 1: Upload
        document_id, document_key = upload_document(source_file, target_lang)
        print(f"   ✅ Document uploaded. ID: {document_id}")

        # Step 2: Wait for translation to complete
        print(f"   ⏳ Waiting for translation to complete...")
        max_attempts = 60  # 5 minutes max
        attempt = 0

        while attempt < max_attempts:
            status = check_translation_status(document_id, document_key)

            if status == 'done':
                print(f"   ✅ Translation complete!")
                break
            elif status == 'error':
                raise Exception("Translation failed with error status")
            elif status in ['translating', 'queued']:
                print(f"   ... Status: {status} (attempt {attempt + 1}/{max_attempts})")
                time.sleep(5)  # Wait 5 seconds before checking again
                attempt += 1
            else:
                raise Exception(f"Unknown status: {status}")

        if attempt >= max_attempts:
            raise Exception("Translation timeout - took too long")

        # Step 3: Download
        download_translated_document(document_id, document_key, output_file)
        print(f"   ✅ Saved to: {output_file}")

        return True

    except Exception as e:
        print(f"   ❌ Error: {e}")
        return False


def main():
    """Main translation function"""
    print("\n" + "="*80)
    print("DeepL Document Translation")
    print("="*80)

    # Check if source file exists
    if not os.path.exists(SOURCE_FILE):
        print(f"\n❌ Error: Source file not found: {SOURCE_FILE}")
        print("   Please make sure the converted EUR document exists.")
        return

    # Create output directory if it doesn't exist
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"\n✅ Source file: {SOURCE_FILE}")
    print(f"✅ Output directory: {OUTPUT_DIR}")
    print(f"✅ Languages to translate: {len(LANGUAGES)}")

    # Translate to each language
    successful = 0
    failed = 0

    for lang_code, lang_name in LANGUAGES.items():
        # Generate output filename
        output_filename = f"month recap_{lang_code}.docx"
        output_path = os.path.join(OUTPUT_DIR, output_filename)

        # Translate
        if translate_document(SOURCE_FILE, lang_code, output_path):
            successful += 1
        else:
            failed += 1

        # Small delay between requests to avoid rate limiting
        time.sleep(2)

    # Summary
    print("\n" + "="*80)
    print("TRANSLATION SUMMARY")
    print("="*80)
    print(f"✅ Successful translations: {successful}/{len(LANGUAGES)}")
    if failed > 0:
        print(f"❌ Failed translations: {failed}/{len(LANGUAGES)}")

    print(f"\nAll translated documents saved to:")
    print(f"  {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
