"""
携程武汉旅游产品爬虫 - 主程序入口
"""
import sys
from crawler import CtripTourCrawler
from datetime import datetime

def main():
    print("=" * 60)
    print("携程武汉旅游产品爬虫")
    print("=" * 60)
    
    crawler = CtripTourCrawler(
        destination="武汉",
        output_file="tours_wuhan.xlsx"
    )
    
    try:
        print("📍 正在爬取携程武汉旅游产品...")
        tours = crawler.crawl_all_tours()
        
        if tours:
            print(f"✅ 共爬取 {len(tours)} 个产品")
            crawler.save_to_excel(tours)
            print(f"✅ Excel 文件已保存: tours_wuhan.xlsx")
        else:
            print("❌ 未爬取到任何产品")
            
    except Exception as e:
        print(f"❌ 爬虫出错: {str(e)}")
    finally:
        crawler.close()

if __name__ == "__main__":
    main()
