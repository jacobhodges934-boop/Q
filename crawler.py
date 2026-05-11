"""
携程旅游产品爬虫核心逻辑 - 适配度假村产品列表
"""

import time
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import logging
import re

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class CtripTourCrawler:
    def __init__(self, destination="武汉", output_file="tours_wuhan.xlsx", base_url=None):
        self.destination = destination
        self.output_file = output_file
        self.base_url = base_url or "https://vacations.ctrip.com/list/whole/sc477.html"
        self.driver = self._init_driver()
        self.tours = []
        self.session = requests.Session()
        self.session.headers.update(self._get_headers())
        
    def _init_driver(self):
        """初始化 Selenium WebDriver"""
        options = Options()
        # 可选：添加 --headless 参数在后台运行
        # options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        
        try:
            # 使用 webdriver-manager 自动管理 ChromeDriver
            driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()),
                options=options
            )
            driver.set_page_load_timeout(30)
            logger.info("✅ Chrome 浏览器已启动")
            return driver
        except Exception as e:
            logger.error(f"❌ 启动 Chrome 失败: {str(e)}")
            raise
    
    @staticmethod
    def _get_headers():
        """获取随机 User-Agent"""
        user_agents = [
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:121.0) Gecko/20100101 Firefox/121.0",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        ]
        return {
            "User-Agent": random.choice(user_agents),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Accept-Encoding": "gzip, deflate",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
            "Referer": "https://www.ctrip.com/",
        }
    
    def _random_delay(self, min_sec=1, max_sec=3):
        """随机延迟"""
        delay = random.uniform(min_sec, max_sec)
        time.sleep(delay)
    
    def crawl_all_tours(self):
        """爬取所有旅游产品"""
        tours = []
        page = 1
        max_pages = 50  # 最多爬取50页
        consecutive_empty = 0
        
        while page <= max_pages:
            logger.info(f"正在爬取第 {page} 页...")
            
            try:
                page_tours = self._crawl_page(page)
                
                if not page_tours:
                    consecutive_empty += 1
                    logger.warning(f"第 {page} 页没有产品 (连续空页: {consecutive_empty})")
                    
                    if consecutive_empty >= 2:
                        logger.info("连续2页未找到产品，停止爬取")
                        break
                else:
                    consecutive_empty = 0
                    tours.extend(page_tours)
                    logger.info(f"第 {page} 页爬取了 {len(page_tours)} 个产品，累计 {len(tours)} 个")
                
                page += 1
                self._random_delay(2, 5)
                
            except Exception as e:
                logger.error(f"爬取第 {page} 页时出错: {str(e)}")
                page += 1
                self._random_delay(3, 6)
        
        logger.info(f"✅ 爬取完成，总共获得 {len(tours)} 个产品")
        return tours
    
    def _crawl_page(self, page):
        """爬取单个页面"""
        try:
            # 构造 URL 带分页参数
            if page == 1:
                url = self.base_url
            else:
                # 添加分页参数
                separator = "&" if "?" in self.base_url else "?"
                url = f"{self.base_url}{separator}pageindex={page}"
            
            logger.info(f"访问 URL: {url}")
            
            # 使用 Selenium 加载页面
            self.driver.get(url)
            
            # 等待产品列表加载
            wait = WebDriverWait(self.driver, 15)
            try:
                # 等待产品容器加载 - 尝试多个选择器
                wait.until(EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "[class*='product'], [class*='item'], li[data-id], .list-item")
                ))
                logger.info("页面产品已加载")
            except:
                logger.warning("页面加载超时")
            
            self._random_delay(1, 3)
            
            # 获取页面源码
            html = self.driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            
            # 查找产品容器 - 根据实际页面结构调整
            tour_items = []
            
            # 尝试多种可能的选择器
            selectors = [
                ('ul.vac-list li', 'list item selector'),
                ('div[class*="product-item"]', 'product-item class'),
                ('div[class*="vacation-item"]', 'vacation-item class'),
                ('li[data-id]', 'data-id attribute'),
                ('div[class*="item-box"]', 'item-box class'),
            ]
            
            for selector, desc in selectors:
                try:
                    elements = soup.select(selector)
                    if elements:
                        tour_items = elements
                        logger.info(f"找到产品容器 ({desc}): {len(elements)} 个")
                        break
                except:
                    continue
            
            if not tour_items:
                logger.warning("未找到任何产品元素，尝试通过 HTML 结构查找...")
                # 最后的尝试 - 查找所有包含特定文本的元素
                tour_items = soup.find_all(['li', 'div'], attrs={'class': lambda x: x and any(
                    key in str(x).lower() for key in ['item', 'product', 'vacation', 'list']
                )})
            
            logger.info(f"找到 {len(tour_items)} 个产品元素")
            
            if len(tour_items) == 0:
                return []
            
            page_tours = []
            for idx, item in enumerate(tour_items[:50], 1):  # 限制每页最多50个产品
                try:
                    tour = self._parse_tour_item(item)
                    if tour and self._is_valid_tour(tour):
                        page_tours.append(tour)
                        logger.debug(f"成功解析第 {idx} 个产品: {tour.get('产品名称', 'N/A')}")
                except Exception as e:
                    logger.debug(f"解析第 {idx} 个产品出错: {str(e)}")
                    continue
            
            logger.info(f"第 {page} 页成功解析了 {len(page_tours)} 个产品")
            return page_tours
            
        except Exception as e:
            logger.error(f"爬取页面出错: {str(e)}")
            return []
    
    def _parse_tour_item(self, item):
        """解析单个产品"""
        try:
            tour = {}
            
            # 产品名称 - 尝试多种方式
            name = 'N/A'
            for selector in ['h3', 'h2', 'a[title]', '[class*="title"]', '[class*="name"]']:
                try:
                    elem = item.select_one(selector)
                    if elem:
                        name = elem.get_text(strip=True)[:100]  # 限制长度
                        break
                except:
                    continue
            tour['产品名称'] = name
            
            # 产品 URL
            link = 'N/A'
            try:
                link_elem = item.find('a', href=True)
                if link_elem:
                    href = link_elem.get('href', '')
                    if href:
                        link = href if href.startswith('http') else f"https://vacations.ctrip.com{href}"
            except:
                pass
            tour['产品链接'] = link
            
            # 价格
            price = 'N/A'
            for price_selector in [
                '[class*="price"]', '[class*="rmb"]', '[class*="cost"]', 
                'span[class*="money"]', 'em[class*="price"]'
            ]:
                try:
                    price_elem = item.select_one(price_selector)
                    if price_elem:
                        price_text = price_elem.get_text(strip=True)
                        price_match = re.search(r'\d+', price_text.replace(',', ''))
                        if price_match:
                            price = price_match.group()
                            break
                except:
                    continue
            tour['价格'] = price
            
            # 天数
            days = 'N/A'
            for days_selector in ['[class*="day"]', '[class*="duration"]', 'span']:
                try:
                    days_elem = item.select_one(days_selector)
                    if days_elem:
                        days_text = days_elem.get_text(strip=True)
                        days_match = re.search(r'(\d+)\s*天', days_text)
                        if days_match:
                            days = days_match.group(1)
                            break
                except:
                    continue
            tour['天数'] = days
            
            # 出发地
            tour['出发地'] = '武汉'
            
            # 目的地
            tour['目的地'] = self.destination
            
            # 评分 - 从星星或数字查找
            score = 'N/A'
            for score_selector in ['[class*="score"]', '[class*="rating"]', '[class*="star"]']:
                try:
                    score_elem = item.select_one(score_selector)
                    if score_elem:
                        score_text = score_elem.get_text(strip=True)
                        score = score_text[:5]  # 取前5个字符
                        break
                except:
                    continue
            tour['评分'] = score
            
            # 评价数
            review_count = '0'
            text_content = item.get_text()
            review_match = re.search(r'(\d+)\s*(?:条|个)?(?:评|点评|评价)', text_content)
            if review_match:
                review_count = review_match.group(1)
            tour['评价数'] = review_count
            
            # 关键词/特色
            tour['关键词'] = self._extract_keywords(tour['产品名称'])
            
            return tour
            
        except Exception as e:
            logger.debug(f"解析产品出错: {str(e)}")
            return None
    
    @staticmethod
    def _extract_keywords(title):
        """从标题中提取关键词"""
        keywords = set()
        
        keyword_map = {
            '温泉': ['温泉'],
            '漂流': ['漂流'],
            '自驾': ['自驾'],
            '跟团': ['跟团', '团队'],
            '蜜月': ['蜜月'],
            '亲子': ['亲子', '亲子游'],
            '爸妈': ['爸妈', '父母'],
            '山水': ['山水', '名山'],
            '古镇': ['古镇', '水乡'],
            '海滨': ['海滨', '海边', '海岛'],
            '高铁': ['高铁'],
            '飞机': ['飞机', '航班'],
            '自由行': ['自由行', '自助'],
            '周末': ['周末'],
            '国际': ['国际', '出国'],
        }
        
        title_lower = title.lower()
        
        for key, patterns in keyword_map.items():
            for pattern in patterns:
                if pattern in title_lower:
                    keywords.add(key)
                    break
        
        return ','.join(keywords) if keywords else '其他'
    
    @staticmethod
    def _is_valid_tour(tour):
        """验证旅游产品数据"""
        required_fields = ['产品名称', '目的地']
        for field in required_fields:
            value = tour.get(field)
            if not value or value == 'N/A' or len(str(value).strip()) == 0:
                return False
        return True
    
    def save_to_excel(self, tours):
        """保存到 Excel 文件"""
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "旅游产品"
            
            # 定义列
            columns = ['产品名称', '价格', '出发地', '目的地', '天数', '评分', '评价数', '关键词', '产品链接']
            
            # 写入表头
            for col_idx, col_name in enumerate(columns, 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.value = col_name
                
                # 设置表头样式
                cell.font = Font(bold=True, color="FFFFFF", size=12)
                cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
            # 定义边框
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # 写入数据
            for row_idx, tour in enumerate(tours, 2):
                for col_idx, col_name in enumerate(columns, 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    value = tour.get(col_name, 'N/A')
                    cell.value = value
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            
            # 调整列宽
            column_widths = {
                'A': 30,  # 产品名称
                'B': 12,  # 价格
                'C': 12,  # 出发地
                'D': 12,  # 目的地
                'E': 10,  # 天数
                'F': 10,  # 评分
                'G': 10,  # 评价数
                'H': 20,  # 关键词
                'I': 35,  # 产品链接
            }
            
            for col, width in column_widths.items():
                ws.column_dimensions[col].width = width
            
            # 设置行高
            ws.row_dimensions[1].height = 25
            for row in range(2, len(tours) + 2):
                ws.row_dimensions[row].height = 20
            
            # 冻结表头
            ws.freeze_panes = "A2"
            
            # 保存文件
            wb.save(self.output_file)
            logger.info(f"✅ Excel 文件已保存: {self.output_file}")
            logger.info(f"📊 共保存 {len(tours)} 个产品")
            
        except Exception as e:
            logger.error(f"保存 Excel 出错: {str(e)}")
            raise
    
    def close(self):
        """关闭浏览器"""
        if self.driver:
            try:
                self.driver.quit()
                logger.info("✅ 浏览器已关闭")
            except Exception as e:
                logger.warning(f"关闭浏览器出错: {str(e)}")
