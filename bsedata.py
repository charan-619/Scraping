from playwright.sync_api import sync_playwright, TimeoutError
import pandas as pd
import time

def scrape_reliance_governance_data(max_retries=3, wait_time=10):
    with sync_playwright() as p:
        for attempt in range(max_retries):
            try:
                browser = p.chromium.launch(headless=False)  # Set to True for production
                page = browser.new_page()
                
                # Navigate to the Reliance Industries page
                page.goto("https://www.bseindia.com/stock-share-price/reliance-industries-ltd/reliance/500325/", wait_until="networkidle")
                
                # Wait for the page to load and click on the Corporate Governance tab
                page.wait_for_selector("text=Corporate Governance", timeout=60000)
                page.click("text=Corporate Governance")
                
                # Wait for content to load
                page.wait_for_load_state("networkidle")
                
                # Extract all text content from the page
                content = page.evaluate("""
                    () => {
                        return document.body.innerText;
                    }
                """)
                
                # Process the content to extract board members
                lines = content.split('\n')
                board_members = []
                capture = False
                for line in lines:
                    if "Board of Directors" in line:
                        capture = True
                        continue
                    if capture and line.strip() == "":
                        capture = False
                        break
                    if capture:
                        board_members.append(line.strip())
                
                browser.close()
                
                if board_members:
                    # Create DataFrame
                    df = pd.DataFrame(board_members, columns=['Board Member'])
                    return df
                else:
                    print("No board member data found. Retrying...")
                    time.sleep(wait_time)
            except TimeoutError:
                print(f"Timeout error occurred. Attempt {attempt + 1} of {max_retries}")
                time.sleep(wait_time)
            except Exception as e:
                print(f"An error occurred: {str(e)}")
                time.sleep(wait_time)
    
    print("Failed to scrape data after multiple attempts")
    return None

def save_to_excel(df, filename):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Board of Directors')
        
        # Auto-adjust columns' width
        for column in df:
            column_width = max(df[column].astype(str).map(len).max(), len(column))
            col_idx = df.columns.get_loc(column)
            writer.sheets['Board of Directors'].column_dimensions[chr(65 + col_idx)].width = column_width + 2

    print(f"Data saved to {filename}")

if __name__ == "__main__":
    data = scrape_reliance_governance_data()
    if data is not None and not data.empty:
        print(data)
        save_to_excel(data, "reliance_governance_data.xlsx")
    else:
        print("Failed to retrieve data or the data is empty.")