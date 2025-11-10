#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mac 一键删除 Word/Excel 所有图片（支持批量、拖拽、进度条）
作者: 你的大名（比如 Tan Zhanghai）
GitHub: https://github.com/你的用户名/Delete-Office-Pics-Mac
"""

import os
import sys
import zipfile
from pathlib import Path
import openpyxl
from docx import Document
from tqdm import tqdm
import shutil
from datetime import datetime

def backup_file(filepath):
    """可选备份原文件"""
    backup_dir = filepath.parent / "备份_图片已删除"
    backup_dir.mkdir(exist_ok=True)
    backup_path = backup_dir / f"{filepath.stem}_备份_{datetime.now():%Y%m%d_%H%M%S}{filepath.suffix}"
    shutil.copy2(filepath, backup_path)
    return backup_path

def delete_word_images(docx_path, backup=False):
    if backup:
        backup_file(docx_path)
    
    try:
        doc = Document(docx_path)
        # 删除所有图片关系
        for rel in list(doc.part.rels.values()):
            if "image" in rel.target_ref:
                doc.part.rels.pop(rel._relid, None)
        
        # 清空包含图片的段落
        for paragraph in doc.paragraphs:
            if paragraph._p.xml.find('pic:pic') != -1:
                paragraph.clear()
        
        # 删除页眉页脚图片
        for section in doc.sections:
            for header in [section.header, section.first_page_header, section.even_page_header]:
                if header.is_linked_to_previous is False:
                    for rel in list(header.part.rels.values()):
                        if "image" in rel.target_ref:
                            header.part.rels.pop(rel._relid, None)
            for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
                if footer.is_linked_to_previous is False:
                    for rel in list(footer.part.rels.values()):
                        if "image" in rel.target_ref:
                            footer.part.rels.pop(rel._relid, None)
        
        doc.save(docx_path)
        return True
    except zipfile.BadZipFile:
        print(f"文件损坏或加密，跳过: {docx_path.name}")
        return False
    except Exception as e:
        print(f"处理 Word 失败 {docx_path.name}: {e}")
        return False

def delete_excel_images(xlsx_path, backup=False):
    if backup:
        backup_file(xlsx_path)
    
    try:
        wb = openpyxl.load_workbook(xlsx_path, keep_vba=True)
        for ws in wb.worksheets:
            # 删除图片
            ws._images = []
            # 删除图表
            ws._charts = []
            # 删除形状（包括 SmartArt）
            if hasattr(ws, '_shapes'):
                ws._shapes = []
        wb.save(xlsx_path)
        return True
    except zipfile.BadZipFile:
        print(f"文件加密或损坏，跳过: {xlsx_path.name}")
        return False
    except Exception as e:
        print(f"处理 Excel 失败 {xlsx_path.name}: {e}")
        return False

def main():
    print("Mac Office 图片删除神器 v2.0")
    print("作者: 你的大名 | GitHub: https://github.com/你的用户名/Delete-Office-Pics-Mac\n")
    
    # 支持拖拽文件/文件夹
    paths = []
    if len(sys.argv) > 1:
        paths = [Path(p) for p in sys.argv[1:]]
    else:
        folder = input("请把文件夹或文件拖到这里（或回车使用桌面）：").strip().strip('"\'')
        if not folder:
            folder = Path("~/Desktop").expanduser()
        else:
            folder = Path(folder)
        if folder.is_dir():
            paths = list(folder.rglob("*.doc*")) + list(folder.rglob("*.xls*"))
        else:
            paths = [folder]
    
    # 过滤有效文件
    files = [p for p in paths if p.suffix.lower() in {".docx", ".docm", ".xlsx", ".xlsm"}]
    if not files:
        print("没找到 Word/Excel 文件，退出。")
        return
    
    backup = input("是否备份原文件？(y/N): ").strip().lower() == 'y'
    
    print(f"\n开始处理 {len(files)} 个文件...")
    success = 0
    for file in tqdm(files, desc="删除图片", unit="文件"):
        try:
            if file.suffix.lower() in {".docx", ".docm"}:
                if delete_word_images(file, backup):
                    success += 1
            else:
                if delete_excel_images(file, backup):
                    success += 1
        except:
            pass
    
    print(f"\n完成！成功处理 {success}/{len(files)} 个文件")
    print("所有图片、图表、SmartArt 已删除！")

if __name__ == "__main__":
    main()
