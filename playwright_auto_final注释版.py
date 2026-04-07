# =================================================================
# 平安银行出账业务自动化提取工具 - 最终完美融合版
# 功能：自动登录、查询借据号、提取7个表格、智能追加到Excel
# =================================================================

# ---------- 标准库导入 ----------
import os                       # 文件和路径操作
import re                       # 正则表达式，用于匹配空单元格
import getpass                  # 安全输入密码（不回显）

# ---------- 第三方库导入 ----------
# type: ignore 消除Pylance缺失类型存根的警告（不影响运行）
from playwright.sync_api import sync_playwright  # type: ignore
import pandas as pd             # 数据处理，Excel读写核心


def _js_extract() -> str:
    """
    【浏览器端JS通用表格提取函数】
    
    工作原理：注入到浏览器执行，直接操作DOM提取表格数据
    返回值：JS函数字符串，供Playwright的evaluate()执行
    
    核心特性：
    1. 兼容模式：自动识别传入的是<table>还是外层<div>
    2. 去重机制：找到第一行（表头）后，跳过后续含RL开头的重复表头行
    3. 去空行：整行全空的行自动过滤
    """
    return """
    (el) => {
        // 兼容层：支持直接传入table或包含table的外层div
        const table = el.querySelector('table') || el;
        
        // 提取所有行元素
        const rows = table.querySelectorAll('tr');
        const data = [];
        
        // 防重复表头标记：找到第一行数据后变为true
        // 后续再遇到RL开头（借据号）的行就是重复表头，自动跳过
        let headerFound = false;

        rows.forEach((row) => {
            // 提取当前行所有单元格文本（自动去前后空格）
            const rowData = [];
            row.querySelectorAll('th, td').forEach(cell => 
                rowData.push(cell.innerText.trim())
            );

            // --------------------------
            // 过滤逻辑：只保留有效数据行
            // --------------------------
            // 1. 跳过纯空行（至少有一个非空单元格才保留）
            if (rowData.some(c => c !== '')) {
                
                // 2. 防重复表头：已找到数据后，遇到RL开头/14位编号的行跳过
                //    （这些是页面上的重复表头分隔行）
                if (rowData.some(c => c.startsWith('RL') || c.length === 14) && headerFound) {
                    return;
                }

                // 3. 有效行：加入结果集
                data.push(rowData);
                headerFound = true;  // 标记：后续RL行是重复表头
            }
        });
        return data;
    }
    """


def main():
    """
    【主流程函数】自动化核心控制逻辑
    
    执行顺序：
    1. 读取待处理的借据号列表
    2. 交互式输入登录账号密码（安全模式）
    3. 启动浏览器完成登录
    4. 逐个借据号循环处理
    """
    # ---------- 初始化配置 ----------
    # 读取借据号：Excel第1列，第2行开始（第1行是表头）
    loan_ids_df = pd.read_excel(r'D:\平安爱码\browser_auto\loan_ids.xlsx', header=None)
    loan_ids = loan_ids_df.iloc[1:, 0].tolist()

    # 创建输出目录（exist_ok=True：已存在不报错）
    output_dir = r'D:\平安爱码\browser_auto\读取结果'
    os.makedirs(output_dir, exist_ok=True)

    # 安全登录：交互式输入，getpass不回显密码
    username = input("请输入用户名: ")
    password = getpass.getpass("请输入密码: ")

    # =================================================================
    # Playwright 浏览器会话上下文
    # =================================================================
    with sync_playwright() as pw:
        # 启动浏览器（headless=False 显示窗口，方便调试）
        browser = pw.chromium.launch(headless=False)
        # 创建上下文（隔离会话，相当于独立浏览器环境）
        context = browser.new_context()
        # 创建新标签页
        page = context.new_page()

        # ---------- 登录流程 ----------
        page.goto("http://ebank.pab.com.cn/bloan/capv-web/#/login")
        page.locator("input[name=\"username\"]").fill(username)
        page.locator("input[name=\"password\"]").fill(password)
        page.get_by_role("button", name="登录").click()

        # ✅ 智能等待：确认登录成功后再操作，避免竞速问题
        #    代替原来的 time.sleep(3)，节省时间且稳定
        page.get_by_text("出账", exact=True).wait_for(state="visible", timeout=15000)
        page.get_by_text("出账", exact=True).click()
        page.get_by_role("menuitem", name="业务查询").click()

        # =================================================================
        # 逐个处理每一个借据号
        # =================================================================
        for loan_id in loan_ids:
            print(f"\n{'='*50}")
            print(f"处理借据号: {loan_id}")
            print(f"{'='*50}")

            try:
                # ---------- 查询借据号 ----------
                page.get_by_role("textbox").first.click()
                page.get_by_role("textbox").first.fill(str(loan_id))
                page.get_by_role("button", name="查询", exact=True).click()

                # ---------- 进入详情页 ----------
                # 点击第4个空单元格（页面约定位置）
                page.get_by_role("cell").filter(has_text=re.compile(r"^$")).nth(3).click()
                
                # 等待详情按钮出现再点击（代替time.sleep(1)）
                page.get_by_role("button", name="查看出账详情").wait_for(
                    state="visible", timeout=10000
                )

                # 捕获新弹出窗口（约定：点击后会开新标签页）
                with page.expect_popup() as page1_info:
                    page.get_by_role("button", name="查看出账详情").click()
                
                detail_page = page1_info.value
                
                # ✅ 等待核心元素出现，代表页面加载完成
                #    代替原来的 time.sleep(3) 硬等待
                detail_page.wait_for_selector(
                    '.el-card__body', state='visible', timeout=15000
                )

                # ---------- 核心：提取所有表格并保存 ----------
                extract_tables(detail_page, output_dir, loan_id)

            except Exception as e:
                print(f"❌ 处理借据号 {loan_id} 时出错: {e}")

            finally:
                # =============================================================
                # 清理机制：无论成功失败，都确保关闭详情页
                # 防止开太多标签页导致内存泄漏
                # =============================================================
                for tab in context.pages:
                    if "putoutDetail" in tab.url:
                        try:
                            tab.close()
                            print(f"✅ 已关闭详情页")
                        except Exception as close_err:
                            print(f"⚠️ 关闭页面失败: {close_err}")
                        
                        # break 只跳出【当前for tab循环】
                        # 找到第一个详情页关闭后就够了，不用继续检查其他页面
                        break  
                else:
                    # for...else 语法：循环正常结束（没被break打断）时执行
                    # = 遍历了所有页面都没找到详情页
                    print("ℹ️ 未找到详情页")

        browser.close()
        print(f"\n{'='*50}")
        print("✅ 所有借据号处理完成！")
        print(f"{'='*50}")


def extract_tables(page, output_dir: str, loan_id) -> None:
    """
    【配置化提取7个表格】
    
    ✅ 设计亮点：以后新增表格只需要在 tables_config 加一行！
    ✅ 7个表格共用同一份JS提取逻辑，代码零重复
    
    参数:
        page: Playwright页面对象（详情页）
        output_dir: 保存Excel的目录
        loan_id: 当前处理的借据号，作为标识列
    """
    # =================================================================
    # 7个表格的配置中心 - 新增表格只需要在这里加元组！
    # 格式: (CSS选择器, 输出Excel文件名)
    # =================================================================
    tables_config = [
        ('.el-card__body',                                      "征信信息(按个人).xlsx"),
        ('.el-card.putout-theThirdPart.credit-score > .el-card__body', "B卡评分.xlsx"),
        ('div:nth-child(4) > .el-card__body',                   "征信信息(按企业).xlsx"),
        ('.el-card__body > div > div > .el-card__body',         "中数工商信息(按企业).xlsx"),
        ('.el-card__body > div > div:nth-child(2) > .el-card__body', "中数工商信息(受托企业).xlsx"),
        ('div:nth-child(2) > div > .el-card__body',             "按自然人查询.xlsx"),
        ('.el-card__body > div:nth-child(2) > .el-card__body',  "按企业查询.xlsx"),
    ]

    # 提前编译好JS提取函数（只编译一次，7个表格共用）
    js_extract = _js_extract()

    # ---------- 循环处理每个表格 ----------
    for selector, filename in tables_config:
        # .first 是统一规范：匹配多个元素时只取第一个，避免strict mode报错
        locator = page.locator(selector).first
        locator.wait_for(state='visible')
        
        # 在浏览器执行JS提取表格数据
        table_data = locator.evaluate(js_extract)
        
        if table_data:
            save_or_append(output_dir, filename, table_data, loan_id)

    print("✅ 所有表格信息提取完成")


def save_or_append(output_dir: str, filename: str, data: list, loan_id) -> None:
    """
    【智能Excel追加存储函数】
    ✅ 彻底解决了原始版本"出账编号隔一列"的核心问题！
    
    设计原理：
        - 放弃原始版本"按位置对齐"的方式
        - 启用 pandas 原生 header=0 机制，【按列名】自动对齐
        - 不管列顺序，不管列数多少，pandas自动匹配
    
    参数:
        output_dir: 输出目录
        filename: Excel文件名
        data: JS提取的二维列表数据 [[表头], [行1], [行2], ...]
        loan_id: 当前借据号，作为每行的标识列
    """
    output_file = os.path.join(output_dir, filename)

    # ---------- Step 1: 新数据清洗 ----------
    df = pd.DataFrame(data)

    # 空数据防护：直接返回不写文件
    if len(df) == 0:
        print(f"⚠️  {filename}: 没有数据可保存")
        return

    # ✅ 第一行设为pandas真正的列名（不再作为数据行存储）
    # ✅ 不再需要关键词识别！任何表头自动适配
    if len(df) >= 1:
        df.columns = df.iloc[0].tolist()   # 第0行 → DataFrame列名
        df = df.iloc[1:].reset_index(drop=True)  # 删掉原来的表头行

    # ---------- Step 2: 添加业务标识列 ----------
    # 所有行统一插入"出账申请编号"，永远在第1列
    df.insert(0, "出账申请编号", loan_id)

    # =================================================================
    # Step 3: 智能追加逻辑 - 核心改进
    # =================================================================
    if os.path.exists(output_file):
        try:
            # ✅ header=0：第一行就是列名，pandas原生支持
            existing_df = pd.read_excel(output_file, header=0)
            
            # ✅ 按列名自动合并！位置无关！
            # 旧文件少列 → pandas自动补NaN
            # 旧文件多列 → 新数据对应列自动补空
            # 列顺序乱 → 自动按列名匹配，绝对不会错位
            combined_df = pd.concat([existing_df, df], ignore_index=True)
            
            # 写入：永远保留header，Excel永远有正确的列名
            combined_df.to_excel(output_file, index=False, header=True)
            print(f"📝  {filename}: 成功追加 {len(df)} 条数据")

        except Exception as e:
            print(f"⚠️  读取旧文件失败，将覆盖写入: {str(e)}")
            df.to_excel(output_file, index=False, header=True)
    else:
        # 文件不存在：直接新建，列名自动写入Excel第一行
        df.to_excel(output_file, index=False, header=True)
        print(f"📄 {filename}: 新建文件，写入 {len(df)} 条数据")


if __name__ == "__main__":
    """程序入口点"""
    main()
