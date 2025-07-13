import openai
import os
from excel import excel_action
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

AZURE_OPENAI_KEY = os.getenv("AZURE_OPENAI_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_OPENAI_DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT")

openai.api_type = "azure"
openai.api_key = AZURE_OPENAI_KEY
openai.api_base = AZURE_OPENAI_ENDPOINT
openai.api_version = "2024-02-15-preview"  # Use the correct API version for your deployment

def get_word_data(word: str) -> dict:
    """
    Given a Norwegian word, use Azure OpenAI to get its English meaning, a Norwegian example sentence,
    and the English translation of that sentence. Returns a dict suitable for excel_action.
    """
    prompt = f"""
    For the Norwegian word: '{word}', provide the following:
    1. The English meaning of the word.
    2. A Norwegian example sentence using the word.
    3. The English translation of that example sentence.
    Respond in JSON with keys: 'word', 'meaning', 'example', 'example_english'.
    """
    client = openai.AzureOpenAI(
        api_key=AZURE_OPENAI_KEY,
        api_version="2024-02-15-preview",
        azure_endpoint=AZURE_OPENAI_ENDPOINT
    )
    response = client.chat.completions.create(
        model=AZURE_OPENAI_DEPLOYMENT,
        messages=[
            {"role": "system", "content": "You are a helpful assistant for language learning."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.2,
        max_tokens=256,
    )
    content = response.choices[0].message.content
    import json
    try:
        data = json.loads(content)
        data['word'] = word
        return data
    except Exception:
        lines = content.split('\n')
        result = {'word': word}
        for line in lines:
            if ':' in line:
                k, v = line.split(':', 1)
                k = k.strip().lower().replace(' ', '_')
                result[k] = v.strip()
        return result

__all__ = ["get_word_data"]
