import scrapy
from scrapy.crawler import CrawlerProcess
import pandas as pd

itemList=[]
class plateScraper(scrapy.Spider):
    name = 'scrapePlate'
    allowed_domains = ['dvlaregistrations.dvla.gov.uk']

    def start_requests(self):
        df=pd.read_excel('data.xlsx')
        columnA_values=df['PLATE']
        for row in columnA_values:
            global  plate_num_xlsx
            plate_num_xlsx=row
            base_url =f"https://dvlaregistrations.dvla.gov.uk/search/results.html?search={plate_num_xlsx}&action=index&pricefrom=0&priceto=&prefixmatches=&currentmatches=&limitprefix=&limitcurrent=&limitauction=&searched=true&openoption=&language=en&prefix2=Search&super=&super_pricefrom=&super_priceto="
            url=base_url
            yield scrapy.Request(url,callback=self.parse, cb_kwargs={'plate_num_xlsx': plate_num_xlsx})

    def parse(self, response, plate_num_xlsx=None):
        plate = response.xpath('//div[@class="resultsstrip"]/a/text()')[1].get()
        price = response.xpath('//div[@class="resultsstrip"]/p/text()')[1].get()
        a = plate.replace(" ", "").strip()

        if plate_num_xlsx==plate.replace(" ","").strip():
            item= {"plate": plate.strip(), "price": price.strip()}
            itemList.append(item)
            yield  item
        else:
            item = {"plate": plate_num_xlsx, "price": "-"}
            itemList.append(item)
            yield item

        with pd.ExcelWriter('output_res.xlsx', mode='r+',if_sheet_exists='overlay') as writer:
            df_output = pd.DataFrame(itemList)
            df_output.to_excel(writer, sheet_name='result', index=False, header=True)

process = CrawlerProcess()
process.crawl(plateScraper)
process.start()