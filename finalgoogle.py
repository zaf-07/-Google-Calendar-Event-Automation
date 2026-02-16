from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import csv
import time
from typing import List, Dict, Optional
from datetime import datetime


class CalendarEventCreator:

    DAY_MAP = {
        "Sunday": "#xRecDay0",
        "Monday": "#xRecDay1",
        "Tuesday": "#xRecDay2",
        "Wednesday": "#xRecDay3",
        "Thursday": "#xRecDay4",
        "Friday": "#xRecDay5",
        "Saturday": "#xRecDay6"
    }

    def __init__(self, csv_file: str):
        self.csv_file = "Classes - Sheet1.csv"
        self.classes = []
        self.successful_events = []
        self.failed_events = []

    def read_classes_from_csv(self) -> List[Dict]:
        try:
            classes = []
            with open(self.csv_file, 'r', encoding='utf-8') as file:
                reader = csv.DictReader(file)
                for row in reader:
                    classes.append(row)

            if not classes:
                raise ValueError("CSV file is empty")

            print(f"✓ Loaded {len(classes)} classes from CSV")
            print(f"  Columns: {list(classes[0].keys())}\n")
            return classes

        except FileNotFoundError:
            print(f"❌ Error: CSV file '{self.csv_file}' not found")
            raise
        except Exception as e:
            print(f"❌ Error reading CSV: {str(e)}")
            raise

    def fill_event_details(self, page, class_data: Dict) -> bool:

        try:
            values = list(class_data.values())
            class_name = values[0] if len(values) > 0 else ''
            code = values[1] if len(values) > 1 else ''
            start_time = values[2] if len(values) > 2 else ''
            end_time = values[3] if len(values) > 3 else ''
            room = values[4] if len(values) > 4 else ''
            date1 = values[5] if len(values) > 5 else ''
            date2 = values[6] if len(values) > 6 else ''

            title = f"{class_name} - {code}".strip(' -')

            if not title:
                print("  ⚠ Warning: Empty title, skipping event")
                return False

            print(f"  → Filling: '{title}'")

            page.fill('input#xTiIn', title)
            time.sleep(0.5)

            page.fill('input#xStDaIn', 'August 15, 2026')
            time.sleep(0.2)
            page.fill('input#xEnDaIn', 'August 15, 2026')
            time.sleep(0.2)

            if start_time:
                print(f"    • Start time: {start_time}")
                page.click('input#xStTiIn')
                time.sleep(0.2)
                page.keyboard.press('Control+A')
                page.keyboard.type(start_time, delay=50)
                page.keyboard.press('Tab')
                time.sleep(0.2)

            if end_time:
                print(f"    • End time: {end_time}")
                page.click('input#xEnTiIn')
                time.sleep(0.2)
                page.keyboard.press('Control+A')
                page.keyboard.type(end_time, delay=50)
                page.keyboard.press('Tab')
                time.sleep(0.2)

            print(f"    • Setting recurrence for: {date1}, {date2}")
            page.click('div[role="combobox"][aria-label="Recurrence"]')
            time.sleep(0.3)
            page.keyboard.press("End")
            time.sleep(0.1)
            page.keyboard.press('Enter')
            time.sleep(0.3)

            if date1 in self.DAY_MAP:
                page.locator(self.DAY_MAP[date1]).click()
                time.sleep(0.1)

            if date2 in self.DAY_MAP:
                page.locator(self.DAY_MAP[date2]).click()
                time.sleep(0.1)

            page.locator(self.DAY_MAP['Saturday']).click()
            time.sleep(0.1)
            #
            # jsname = "YPqjbf"
            # jsname = "m9ZlFb"
            # Click "On" radio button
            page.locator('label:has-text("On")').click()
            page.keyboard.press('Tab')
            page.keyboard.type('December 15, 2026')

            page.get_by_role("button", name="Done").click()
            time.sleep(0.3)

            if room:
                print(f"    • Adding room: {room}")
                page.locator("#xRmTab").click()
                time.sleep(0.3)

                rooms_panel = page.locator('div[role="tabpanel"][aria-labelledby="xRmTab"]')
                search_box = rooms_panel.locator('div[jsname="vhZMvf"]')
                search_box.wait_for(state="visible", timeout=5000)
                search_box.click()
                time.sleep(0.2)

                search_box.press_sequentially(room, delay=50)
                time.sleep(0.8)

                results_list = page.locator('ul[role="listbox"][aria-label="Search room or resource"]')
                results_list.wait_for(state="visible", timeout=5000)
                time.sleep(0.3)

                second_result = results_list.locator('li, [role="option"]').nth(1)
                second_result.click()
                time.sleep(0.5)

            print(f"    • Removing default guest")
            page.locator("#xGstTab").click()
            time.sleep(0.3)

            guests_panel = page.locator('div[role="tabpanel"][aria-labelledby="xGstTab"]')
            remove_button = guests_panel.get_by_role("button", name="Remove Zafar Azatov from")
            time.sleep(0.5)
            remove_button.dispatch_event("click")
            time.sleep(0.3)

            page.click('input[aria-label="Add location"]')
            page.keyboard.press('Tab')
            page.keyboard.press('Tab')
            page.keyboard.press('Tab')
            page.keyboard.press('Tab')
            page.keyboard.press('Tab')
            page.keyboard.press('Tab')
            page.keyboard.press('Enter')
            time.sleep(0.2)
            page.keyboard.press('ArrowDown')
            page.keyboard.press('Enter')


            print(f"    • Saving event...")
            save_button = page.get_by_role("button", name="Save")
            save_button.click()

            time.sleep(1.5)

            try:
                page.wait_for_selector('[role="dialog"]', state='detached', timeout=8000)
                print(f"    ✓ Event saved successfully")
                time.sleep(0.5)
                return True
            except PlaywrightTimeoutError:
                print(f"    ⚠ Dialog didn't close, but may have saved")
                page.keyboard.press('Escape')
                time.sleep(1)
                return True

        except Exception as e:
            print(f"    ❌ Error filling event: {str(e)}")
            return False

    def create_single_event(self, page, class_data: Dict, event_number: int, total: int) -> bool:

        class_name = class_data.get('Class', 'Unknown')
        print(f"\n[{event_number}/{total}] Creating: {class_name}")

        try:
            create_button = page.locator('button:has-text("Create")')
            create_button.wait_for(state='visible', timeout=10000)
            create_button.click()
            time.sleep(1)

            page.keyboard.press('ArrowUp')
            time.sleep(0.1)
            page.keyboard.press('Enter')
            time.sleep(1)

            page.wait_for_selector('[role="dialog"]', state='visible', timeout=10000)
            time.sleep(0.3)

            more_options = page.locator('button:has-text("More options")')
            more_options.click()
            time.sleep(2)

            page.wait_for_selector('input#xTiIn', state='visible', timeout=10000)
            time.sleep(0.3)

            success = self.fill_event_details(page, class_data)

            if success:
                self.successful_events.append(class_data)
                print(f"✓ Event #{event_number} completed: {class_name}")
            else:
                self.failed_events.append(class_data)
                print(f"⚠ Event #{event_number} had issues: {class_name}")

            return success

        except Exception as e:
            print(f"❌ Failed to create event #{event_number}: {str(e)}")
            self.failed_events.append(class_data)

            try:
                page.keyboard.press('Escape')
                time.sleep(0.5)
                page.keyboard.press('Escape')
                time.sleep(0.5)
            except:
                pass

            return False

    def run(self):

        print("=" * 70)
        print("Google Calendar Event Creator".center(70))
        print("=" * 70)
        print()

        try:
            self.classes = self.read_classes_from_csv()
        except Exception:
            return

        with sync_playwright() as p:
            print("→ Launching browser...")
            context = p.chromium.launch_persistent_context(
                user_data_dir='./chrome-data',
                headless=False,
                channel='chrome',
                args=[
                    '--disable-blink-features=AutomationControlled',
                    '--no-first-run',
                    '--no-default-browser-check',
                ]
            )

            page = context.pages[0] if context.pages else context.new_page()
            page.set_default_timeout(15000)

            print("→ Navigating to Google Calendar...")
            page.goto('https://calendar.google.com', wait_until='domcontentloaded')
            time.sleep(8)

            try:
                page.wait_for_selector('button:has-text("Create")', state='visible', timeout=20000)
                print("✓ Calendar loaded successfully\n")
            except PlaywrightTimeoutError:
                print("❌ Error: Google Calendar didn't load. Please check your login.")
                input("\nPress Enter to close...")
                return

            print("=" * 70)
            print(f"Processing {len(self.classes)} events...".center(70))
            print("=" * 70)

            for i, class_data in enumerate(self.classes, 1):
                self.create_single_event(page, class_data, i, len(self.classes))

                if i < len(self.classes):
                    time.sleep(2)

            print("\n" + "=" * 70)
            print("SUMMARY".center(70))
            print("=" * 70)
            print(f"✓ Successful: {len(self.successful_events)}/{len(self.classes)}")
            print(f"❌ Failed: {len(self.failed_events)}/{len(self.classes)}")

            if self.failed_events:
                print("\nFailed events:")
                for event in self.failed_events:
                    print(f"  - {event.get('Class', 'Unknown')}")

            print("=" * 70)

            input("\nPress Enter to close browser...")


def main():
    """Entry point."""
    csv_file = 'Classes_-_Sheet1.csv'

    creator = CalendarEventCreator(csv_file)
    creator.run()


if __name__ == "__main__":
    main()