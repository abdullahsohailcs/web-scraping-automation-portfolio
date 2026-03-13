
import asyncio, re, os, copy, random
from playwright.async_api import async_playwright
from datetime import datetime

OUTPUT_DIR = r"C:\Users\HP\Desktop"
present_date = datetime.today().strftime('%Y-%m-%d')
portugal_file = os.path.join(OUTPUT_DIR, f"Portugal Output LL {present_date}.xlsx")
translated_file = os.path.join(OUTPUT_DIR, f"Portugal Output Translated {present_date}.xlsx")
bug_file = os.path.join(OUTPUT_DIR, f"Portugal Bugs {present_date}.xlsx")


async def safe_get_text(page, selector):
    try:
        element = await page.query_selector(selector)
        if element:
            return (await element.inner_text()).strip()
        return ""
    finally:
        return ""


async def simulate_human_interaction(page):
    # Random mouse movements
    await page.mouse.move(
        random.randint(0, 800),
        random.randint(0, 600)
    )
    # Random scrolling
    if random.random() > 0.5:
        await page.mouse.wheel(0, random.randint(100, 300))

async def random_delay(min=1000, max=3000):
    await asyncio.sleep(random.randint(min, max) / 1000)




async def run():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()
        page.set_default_timeout(60000)

        await page.goto("https://extranet.infarmed.pt/INFOMED-fo/pesquisa-avancada.xhtml")
        await page.wait_for_selector("div.card-body div.form-group")

        # Apply filter: Temporariamente indisponível
        await page.click("#mainForm\\:estado-comercializacao")
        await page.click("li[data-label='Temporariamente indisponível']")

        # Click search button
        await page.click("#mainForm\\:btnDoSearch")

        # Wait for results to load
        await page.wait_for_selector(".ui-datatable-data", state="visible", timeout=30000)
        await page.wait_for_timeout(5000)

        # Check if results exist
        if await page.query_selector(".ui-datatable-empty-message"):
            print("No results found")
            await browser.close()
            return

        results = []
        bug_results = []
        ee_results = []
        try:
            while True:
                # Wait for results to load
                await page.wait_for_selector(".ui-datatable-data", state="visible", timeout=30000)
                print("==========Table appers=========")
                await page.wait_for_timeout(5000)
                print("==========Table appers=========")

                # Check if results exist
                if await page.query_selector(".ui-datatable-empty-message"):
                    print("No results found")
                    await browser.close()
                    return

                for i in range(10):
                    data = {}
                    bug_data = {}
                    try:

                        # Re-select links fresh each time (DOM updates after AJAX)
                        links = await page.query_selector_all('td[role="gridcell"] div .ui-commandlink')
                        # Get all rows inside tbody where role = row
                        rows = await page.query_selector_all("tbody tr[role='row']")
                        if i >= len(links):
                            break

                        td4_elem = await rows[i].query_selector("td:nth-of-type(4)")
                        dosage = (await td4_elem.inner_text()).strip() if td4_elem else ""
                        data["Product Strength"] = dosage

                        td5_elem = await rows[i].query_selector("td:nth-of-type(5)")
                        dosage = (await td5_elem.inner_text()).strip() if td5_elem else ""
                        data["Product Strength"] = dosage
                        print(data)

                        # Get 6th td text
                        td6_elem = await rows[i].query_selector("td:nth-of-type(6)")
                        company = (await td6_elem.inner_text()).strip() if td6_elem else ""
                        data["Corporation"] = company

                        # Click the link to load detail via AJAX
                        await links[i].click()
                        await page.wait_for_timeout(5000)

                        # Wait for detail view panel
                        await page.wait_for_selector("#panel-detalhes", state="visible", timeout=60000)

                        atc_labels = await page.query_selector_all("#atcId_content label")
                        atc_values = [await label.inner_text() for label in atc_labels]
                        atc_values = [v.strip() for v in atc_values if v.strip()]
                        atc_combined = "; ".join(atc_values)

                        product_name_elem = await page.query_selector("div#pageTitleDetalhe h1 strong")
                        product_name = (await product_name_elem.inner_text()).strip() if product_name_elem else ""

                        # Extract data (your existing safe_get_text calls)
                        scraped_fields = {
                            "Nome do Medicamento": product_name,
                            "ATC": atc_combined,
                        }

                        atc_value = scraped_fields["ATC"]
                        if " - " in atc_value:
                            try:
                                name_part = atc_value.split(" - ", 1)[1]
                                names = [name.strip() for name in name_part.split(" and ")]
                                data["Molecule"] = ";".join(names)
                            except Exception as e:
                                print(f"[Molecule parse error] ATC value: {atc_value} | Error: {e}")
                                data["Molecule"] = ""
                        else:
                            print(f"[Missing dash in ATC] Raw ATC value: '{atc_value}'")
                            data["Molecule"] = ""

                        data["Product Name"] = scraped_fields["Nome do Medicamento"]
                        data["Shortage Status"] = "Temporariamente indisponível"
                        data["Shortage impact rating"] = ""
                        data["End date"] = ""
                        data["Relevant unit number "] = ""
                        data["Additional commentary"] = ""

                        # Check if .alertas-panel exists
                        alertas_elem = await page.query_selector(".alertas-panel")

                        if alertas_elem:
                            # Get the text content
                            alertas_text = (await alertas_elem.inner_text()).strip()

                            if alertas_text:
                                data["Alternatives available"] = "Yes"
                                data["Additional commentary"] = alertas_text
                            else:
                                data["Alternatives available"] = "No"
                        else:
                            data["Alternatives available"] = "No"

                        # Get all matching span elements inside active carousel item
                        span_elems = await page.query_selector_all(
                            "form#carousel-tablet div.carousel-item.active div div span")

                        # Check if there's a second span
                        if len(span_elems) >= 2:
                            span_text = (await span_elems[1].inner_text()).strip()
                        else:
                            span_text = ""
                        data["Product Pack Size"] = span_text

                        # Get the span element text
                        span_elem = await page.query_selector(
                            "form#carousel-tablet div.carousel-item.active div div span.text-card-header")
                        span_text = (await span_elem.inner_text()).strip() if span_elem else ""

                        # Use regex to extract date (format: DD/MM/YYYY)
                        match = re.search(r'\b\d{2}/\d{2}/\d{4}\b', span_text)

                        if match:
                            extracted_date = match.group(0)
                        else:
                            extracted_date = ""
                        data["End date"] = extracted_date
                        data["Date reported"] = present_date
                        data["Date last updated"] = present_date
                        data["Start date"] = present_date

                        results.append(data)
                        ee_data = copy.deepcopy(data)
                        ee_results.append(ee_data)
                        print(data)
                        print(f"Processed medicine {i + 1}")

                        await page.go_back()
                        await page.wait_for_selector(".ui-datatable-data", state="visible", timeout=15000)


                    except Exception as e:
                        print(f"Error processing medicine {i + 1}: {e}")
                        bug_data = copy.deepcopy(data)
                        bug_data["reason"] = str(e)
                        bug_results.append(bug_data)
                        await page.go_back()
                        await page.wait_for_selector(".ui-datatable-data", state="visible", timeout=15000)

                await page.wait_for_timeout(5000)
                next_button = await page.query_selector("div a.ui-paginator-next")
                if next_button:
                    await next_button.click()
                    await page.wait_for_timeout(3000)
                else:
                    break

        finally:
            await browser.close()


asyncio.run(run())