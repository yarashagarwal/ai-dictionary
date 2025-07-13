from getaidata import get_word_data as get_word_data_func
from excel import excel_action
from tabulate import tabulate
import sys
import textwrap

def wrap_entry(entry, width=20):
    return {k: textwrap.fill(str(v), width=width) if isinstance(v, str) else v for k, v in entry.items()}

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Norwegian Dictionary CLI")
    parser.add_argument("-w", "--write", action="store_true", help="Write new words to the Excel dictionary")
    parser.add_argument("-r", "--read", action="store_true", help="Read and display all words in the Excel dictionary")
    parser.add_argument("-s", "--search", type=str, help="Search for a word in the Excel dictionary")
    args = parser.parse_args()

    if args.read:
        entries = excel_action('read')
        if entries:
            print("\nAll words in the Excel dictionary:")
            print(tabulate(entries, headers="keys", tablefmt="grid"))
            print(f"Total words in the Excel dictionary: {len(entries)}")
        else:
            print("The Excel dictionary is empty.")
        sys.exit(0)

    if args.search:
        entries = excel_action('read')
        search_word = args.search.strip().lower()
        found = [entry for entry in entries if entry.get('word', '').lower() == search_word]
        if found:
            print(f"\nSearch result for '{args.search}':")
            print(tabulate(found, headers="keys", tablefmt="grid"))
        else:
            print(f"'{args.search}' not found in the Excel dictionary.")
        sys.exit(0)

    if args.write:
        print("Enter Norwegian words, one per line. Press Enter on an empty line to finish:")
        words = []
        while True:
            word = input().strip()
            if not word:
                break
            words.append(word)
        # Read existing words from Excel before calling the AI API
        existing_entries = excel_action('read')
        existing_words = {entry['word'] for entry in existing_entries if 'word' in entry}
        results = []
        for word in words:
            if word in existing_words:
                print(f"{word}: Already exists in the Excel dictionary. Skipping.")
                continue
            result = get_word_data_func(word)
            print(f"{word}: {result}")
            results.append(result)
        if results:
            # Show results in table format before writing to Excel
            print("\nResults to be added to Excel:")
            wrapped_results = [wrap_entry(entry) for entry in results]
            print(tabulate(wrapped_results, headers="keys", tablefmt="grid"))
            excel_action('append', data=results)
            print(f"Added {len(results)} new entries to the Excel dictionary.")
            # Also append to dict.md
            with open("dict.md", "a", encoding="utf-8") as mdfile:
                for entry in results:
                    mdfile.write(f"\n### {entry.get('word','')}: {entry.get('meaning','')}```\n")
                    mdfile.write(f"- **Example (NO):** {entry.get('example','')}\n")
                    mdfile.write(f"- **Example (EN):** {entry.get('example_english','')}\n")
            print(f"Added {len(results)} new entries to dict.md.")
        else:
            print("No new words to add. All words already exist in the Excel dictionary.")
        print(f"Total words in the Excel dictionary: {len(existing_words) + len(results)}")

        # Reorganize all words alphabetically in the Excel file
        all_entries = excel_action('read')
        sorted_entries = sorted(all_entries, key=lambda x: x.get('word', '').lower())
        excel_action('write', data=sorted_entries)
        print("Reorganized all words alphabetically in the Excel dictionary.")
    else:
        parser.print_help()