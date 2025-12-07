# -*- coding: utf-8 -*-
from DrissionPage import ChromiumPage, ChromiumOptions
import pandas as pd
import time
import random
import os
from urllib.parse import urljoin

# ================= 配置区域 =================
TARGET_URL = "https://ningbo.chinatax.gov.cn/zcwj/zcfgk/index.html"
VERSION = "v12.0 (极速标题版 - 拒绝无效等待)"


def get_desktop_path():
    return os.path.join(os.path.expanduser("~"), "Desktop")


OUTPUT_FILE = os.path.join(get_desktop_path(), "宁波税务_政策法规库_全量抓取.xlsx")


# ================= 核心逻辑 =================

def extract_detail(tab):
    """
    进入详情页后，同时提取：完整标题、元数据、正文、附件
    """
    try:
        info = {
            "标题": "", "正文": "", "文号": "", "发文单位": "", "发布日期": "", "附件": []
        }

        # === 🌟 核心提速优化：限制查找时间 ===
        try:
            # 只给 0.2 秒的时间找标题，找不到立刻换下一个策略
            title_ele = tab.ele('tag:h1', timeout=0.2)

            if not title_ele:
                title_ele = tab.ele('.title', timeout=0.2)

            if not title_ele:
                title_ele = tab.ele('#title', timeout=0.2)

            if title_ele:
                info["标题"] = title_ele.text.strip()
        except:
            pass

        # === 1. Meta数据 ===
        try:
            # Meta 数据通常在头部，不需要等待
            date_ele = tab.ele('xpath://meta[@name="PubDate"]', timeout=0.2)
            if date_ele: info["发布日期"] = date_ele.attr("content").split(" ")[0]
            source_ele = tab.ele('xpath://meta[@name="ContentSource"]', timeout=0.2)
            if source_ele: info["发文单位"] = source_ele.attr("content")
        except:
            pass

        # === 2. 正文 (这是必须存在的，可以多等一会确保加载) ===
        content_ele = tab.ele('#zoom', timeout=5)
        if content_ele:
            info["正文"] = content_ele.text
        else:
            info["正文"] = tab.ele('.info-cont').text if tab.ele('.info-cont') else "正文提取失败"

        # === 3. 文号补救 ===
        if not info["文号"]:
            first_part = info["正文"][:300]
            if "发布文号" in first_part:
                try:
                    parts = first_part.split("发布文号")
                    candidate = parts[1].split("\n")[0].replace("】", "").replace(":", "").replace("：", "").strip()
                    info["文号"] = candidate
                except:
                    pass

        # === 4. 附件 (快速扫描) ===
        # 不需要 wait，直接获取当前已加载的
        links = tab.eles('tag:a')
        for link in links:
            href = link.attr('href')
            if not href: continue
            if href.endswith(('.doc', '.docx', '.xls', '.xlsx', '.pdf', '.zip', '.rar')):
                full_url = urljoin(tab.url, href)
                info["附件"].append({
                    "文件名": link.text,
                    "链接": full_url
                })
        return info

    except Exception as e:
        print(f"    ❌ 详情页解析出错: {e}")
        return {}


def save_to_excel(data_list, filepath):
    if not data_list: return
    while True:
        try:
            df_new = pd.DataFrame(data_list)
            if os.path.exists(filepath):
                try:
                    with pd.ExcelWriter(filepath, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                        pass
                    df_old = pd.read_excel(filepath, engine="openpyxl")
                    df = pd.concat([df_old, df_new], ignore_index=True)
                    df.drop_duplicates(subset=["链接", "附件链接"], keep="last", inplace=True)
                except PermissionError:
                    raise PermissionError
                except:
                    df = df_new
            else:
                df = df_new

            cols = ["标题", "发布日期", "发文单位", "文号", "正文", "附件文件名", "附件链接", "链接"]
            for c in cols:
                if c not in df.columns: df[c] = ""
            df = df[cols]
            df.to_excel(filepath, index=False, engine="openpyxl")
            print(f"   💾 已保存 (总行数: {len(df)})")
            break
        except PermissionError:
            print("\n🚨 错误：Excel 文件被占用！请关闭文件...")
            time.sleep(5)
        except Exception as e:
            print(f"   ❌ Excel保存未知失败: {e}")
            break


def main():
    print(f"🚀 启动采集器 - {VERSION}")

    co = ChromiumOptions()
    co.set_user_agent(
        user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    # 保持禁图，追求极致速度
    co.set_argument('--blink-settings=imagesEnabled=false')
    co.set_argument('--mute-audio')
    co.set_argument('--window-position=-3000,-3000')  # 移出屏幕
    co.ignore_certificate_errors()

    page = ChromiumPage(addr_or_opts=co)

    print(f"🌐 正在访问: {TARGET_URL}")
    page.get(TARGET_URL)
    time.sleep(2)

    processed_urls = set()
    if os.path.exists(OUTPUT_FILE):
        try:
            try:
                df = pd.read_excel(OUTPUT_FILE, engine="openpyxl")
                processed_urls = set(df["链接"].dropna().tolist())
                print(f"📚 已读取 {len(processed_urls)} 条历史记录")
            except:
                pass
        except:
            pass

    page_num = 1
    empty_page_count = 0

    while True:
        print(f"\n🔄 正在处理第 {page_num} 页...")

        try:
            page.wait.ele('tag:a', timeout=8)
        except:
            pass

        all_links = page.eles('tag:a')
        article_links = []
        for link in all_links:
            url = link.attr('href')
            list_title = link.text

            if not url or "javascript" in url: continue
            if not list_title or len(list_title) < 5: continue

            is_article = ("/art/" in url) or ("/content/" in url) or ("202" in url)
            is_category = url.endswith("index.html")

            if is_article and not is_category:
                if url not in processed_urls:
                    article_links.append({"title": list_title, "url": url})

        unique_links = []
        seen = set()
        for item in article_links:
            if item['url'] not in seen:
                unique_links.append(item)
                seen.add(item['url'])

        if not unique_links:
            print("⚠️ 本页未发现新数据。")
            empty_page_count += 1
            if empty_page_count >= 3:
                print("🛑 连续 3 页无数据，判断为结束。")
                break
        else:
            print(f"   📄 筛选出 {len(unique_links)} 篇新文章")
            empty_page_count = 0

        # === 抓取循环 ===
        for item in unique_links:
            short_title = item['title']
            print(f"   Downloading: {short_title[:15]}...")

            try:
                new_tab = page.new_tab(item["url"])
                # 等待正文加载 (这是唯一需要花时间等的)
                new_tab.ele('#zoom', timeout=8)

                detail = extract_detail(new_tab)
                new_tab.close()

                final_title = detail.get("标题")
                if not final_title:
                    final_title = short_title

                row_base = {
                    "标题": final_title,
                    "链接": item["url"],
                    "发布日期": detail.get("发布日期", ""),
                    "发文单位": detail.get("发文单位", ""),
                    "文号": detail.get("文号", ""),
                    "正文": detail.get("正文", "")
                }

                current_data = []
                if detail["附件"]:
                    for att in detail["附件"]:
                        row = row_base.copy()
                        row["附件文件名"] = att["文件名"]
                        row["附件链接"] = att["链接"]
                        current_data.append(row)
                else:
                    row_base["附件文件名"] = ""
                    row_base["附件链接"] = ""
                    current_data.append(row_base)

                processed_urls.add(item["url"])
                save_to_excel(current_data, OUTPUT_FILE)
                # 几乎无延迟的连续抓取
                time.sleep(0.01)
            except Exception as e:
                print(f"   ❌: {e}")
                if page.tabs_count > 1: page.close_tabs(page.tab_ids[1:])

                # === 翻页 (加速版) ===
        print("👆 翻页中...")
        try:
            right_box = page.ele('.right-box')
            if right_box:
                next_btn = right_box.ele('.layui-laypage-next')
            else:
                next_btn = page.ele('.layui-laypage-next')

            if next_btn:
                class_val = next_btn.attr("class")
                if class_val and "disabled" in class_val:
                    print(f"🛑 按钮变灰，结束 (共 {page_num} 页)")
                    break

                next_btn.click(by_js=True)

                # 🌟 从 3秒 缩短到 1.5秒，足够网页刷新了
                print("   ⏳ 等待刷新 (1.5s)...")
                time.sleep(1.5)

                page_num += 1
            else:
                print("🛑 未找到翻页按钮，结束。")
                break

        except Exception as e:
            print(f"🛑 翻页流程出错: {e}")
            break

    print(f"\n🎉 完成！文件: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()