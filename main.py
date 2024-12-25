from scraper import SalarySlipScraper

sss = SalarySlipScraper()

if sss.displays_log_in_page():
    
    sss.import_csv()
    
    sss.try_auto_log_in()
    
    while not sss.completes_log_in():
        sss.prompt_manual_log_in()
    
    sss.close_note()
    
    sss.scrape_slips()
    
    sss.export_csv()
    
sss = None