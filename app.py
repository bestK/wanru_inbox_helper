import pandas as pd
import streamlit as st
import io
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
import zipfile
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
)
import re

st.set_page_config(page_title="Wanru Inbox Helper", page_icon="📬", layout="wide")


def create_pdf(data_df, box_number, box_index, box_count):
    buffer = io.BytesIO()

    # 页面尺寸
    page_width = 100 * mm
    page_height = 150 * mm
    margins = 5 * mm
    usable_width = page_width - 2 * margins
    usable_height = page_height - 2 * margins

    # PDF 文档
    doc = SimpleDocTemplate(
        buffer,
        pagesize=(page_width, page_height),
        rightMargin=margins,
        leftMargin=margins,
        topMargin=margins,
        bottomMargin=margins,
    )

    # 样式
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "CustomTitle",
        parent=styles["Normal"],
        fontSize=10,
        alignment=TA_CENTER,
        spaceAfter=2 * mm,
    )

    # 只选择需要的列，并重命名
    data_df = data_df[["sku", "quantity", "stock_code"]]
    data_df.columns = ["SKU", "QTY", "Stock Code"]
    data = [data_df.columns.tolist()] + data_df.values.tolist()

    # 固定字体大小为14
    font_size = 14
    col_count = len(data[0])
    col_width = usable_width / col_count
    table = Table(data, colWidths=[col_width] * col_count)
    style = TableStyle(
        [
            ("ALIGN", (0, 0), (-2, 0), "CENTER"),
            ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
            ("FONTNAME", (0, 0), (-2, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-2, 0), 18),
            # 库位号单元格大号加粗
            ("FONTNAME", (-1, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (-1, 0), (-1, 0), 50),
            ("ALIGN", (-1, 0), (-1, 0), "CENTER"),
            ("VALIGN", (-1, 0), (-1, 0), "MIDDLE"),
            # 可选：加边框
            ("BOX", (0, 0), (-1, 0), 1, colors.black),
            ("INNERGRID", (0, 0), (-1, 0), 0.5, colors.black),
        ]
    )
    table.setStyle(style)

    # 组装最终的 PDF 内容
    elements = []
    box_info = f"{box_number} ({box_index}/{box_count})"
    elements.append(Paragraph(box_info, title_style))
    elements.append(table)
    doc.build(elements)

    pdf_data = buffer.getvalue()
    buffer.close()
    return pdf_data


def create_sku_pdf(sku, quantity, stock_code, box_number, box_index, box_count):
    buffer = io.BytesIO()

    # 页面设置：竖版
    page_width = 100 * mm
    page_height = 150 * mm
    margins = 5 * mm
    usable_width = page_width - 2 * margins

    doc = SimpleDocTemplate(
        buffer,
        pagesize=(page_width, page_height),
        rightMargin=margins,
        leftMargin=margins,
        topMargin=margins,
        bottomMargin=margins,
    )

    # 样式定义
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "Title",
        parent=styles["Normal"],
        fontSize=10,
        alignment=TA_CENTER,
        spaceAfter=5,
    )

    box_info = f"{box_number} ({box_index}/{box_count})"

    # 表格数据
    data = [["SKU", "QTY", "Stock Code"], [sku, quantity, stock_code]]
    col_count = len(data[0])
    col_width = usable_width / col_count

    table = Table(data, colWidths=[col_width] * col_count)
    table.setStyle(
        TableStyle(
            [
                # ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                # ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 12),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
                # ("BACKGROUND", (0, 1), (-1, -1), colors.beige),
                # ("TEXTCOLOR", (0, 1), (-1, -1), colors.black),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 1), (-1, -1), 12),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ]
        )
    )

    # 库位号样式（放大显示）
    big_style = ParagraphStyle(
        name="BigStockCode",
        fontName="Helvetica-Bold",
        fontSize=100,  # 更大字体
        alignment=TA_CENTER,
        leading=90,
        spaceBefore=40,
    )

    def split_stock_code(stock_code):
        # 匹配前面的字母和后面的数字
        match = re.match(r"([A-Za-z]+)([0-9]+)", str(stock_code))
        if match:
            return f"{match.group(1)}<br/>{match.group(2)}"
        else:
            return str(stock_code)

    # 内容元素按纵向顺序排布
    elements = [
        Paragraph(box_info, title_style),
        Spacer(1, 8),
        table,
        Spacer(1, 20),
        Paragraph(split_stock_code(stock_code), big_style),
    ]

    doc.build(elements)
    pdf_data = buffer.getvalue()
    buffer.close()
    return pdf_data


def main():
    st.title("wanru inbox helper")
    st.write("Welcome to wanru inbox helper!")
    #  please upload box number and SKU excel
    box_excel = st.file_uploader(
        "Please upload [Box number and SKU] excel", type=["xlsx"]
    )

    box_sku_dicts = []
    if box_excel:
        st.write("box number: ", box_excel)
        #  please upload box number and SKU excel
        box_excel_data = pd.read_excel(box_excel)
        box_count = len(box_excel_data)
        st.write("box_count: ", box_count)
        box_skus_str = box_excel_data["Product SKU"].str.split(";")
        # 20 x 4669059408;9 x 8866336620 to [{sku: 4669059408, quantity: 20}, {sku: 8866336620, quantity: 9}]

        for box_index, sku_list in enumerate(box_skus_str):
            box_index_str = str(box_index + 1)
            for sku_str in sku_list:
                parts = sku_str.strip().split(" x ")
                if len(parts) == 2:
                    sku_dict = {
                        "sku": parts[1].strip(),
                        "quantity": int(parts[0]),
                        "box_index": box_index_str,
                        "box_number": box_excel_data["Warehouse receipt code"].iloc[
                            box_index
                        ],
                        "stock_code": "",
                    }
                    box_sku_dicts.append(sku_dict)

        # only write skus from box_sku_dicts to new array
        # eg:4669059408
        # 8866336620
        box_skus = []
        for box_sku_dict in box_sku_dicts:
            if box_sku_dict["sku"] not in box_skus:
                box_skus.append(box_sku_dict["sku"])

        # 将 SKU 数组转换为换行字符串
        skus_text = "\n".join(box_skus)
        st.write("sku type count: ", len(box_skus))
        st.text_area("SKUs:", skus_text, height=300)

        #  please upload Inventory List excel
        inventory_list_excel = st.file_uploader(
            "Please upload [Inventory List] excel", type=["xlsx"]
        )

        # 使用字典存储 inventory_list，key 为 product_sku，value 为 stock_code
        inventory_map = {}
        if inventory_list_excel:
            st.write("inventory_list_excel: ", inventory_list_excel)
            inventory_list_excel_data = pd.read_excel(inventory_list_excel)
            # Stock Code, Product Sku

            # 构建 SKU 到 stock_code 的映射
            for stock_code, product_sku in zip(
                inventory_list_excel_data["Stock Code"],
                inventory_list_excel_data["Product Sku"],
            ):
                inventory_map[str(product_sku).strip()] = str(stock_code).strip()

            # 使用映射更新 box_sku_dicts 中的 stock_code
            for item in box_sku_dicts:
                if item["sku"] in inventory_map:
                    item["stock_code"] = inventory_map[item["sku"]]
                else:
                    st.write(f"SKU not found in inventory: {item['sku']}")

            # 将数据转换为 DataFrame
            box_sku_dicts_df = pd.DataFrame(box_sku_dicts)

            # 创建 Excel 文件的字节流
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                box_sku_dicts_df.to_excel(writer, index=False, sheet_name="Sheet1")

            # 获取 Excel 文件的字节数据
            excel_data = output.getvalue()

            # 添加下载按钮
            st.download_button(
                label="Download processed Excel file",
                data=excel_data,
                file_name="box_sku_dicts.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # 为每个箱号生成 PDF
            pdf_files = []
            for i, box_number in enumerate(box_sku_dicts_df["box_number"].unique(), 1):
                box_sku_dicts_df_box = box_sku_dicts_df[
                    box_sku_dicts_df["box_number"] == box_number
                ]
                st.write(f"Box {box_number} contents:")
                st.dataframe(box_sku_dicts_df_box)

                # 生成 PDF
                pdf_data = create_pdf(box_sku_dicts_df_box, box_number, i, box_count)
                pdf_filename = f"{box_number}.pdf"
                pdf_files.append((pdf_filename, pdf_data))

                # 添加 PDF 下载按钮（单个）
                st.download_button(
                    label=f"Download PDF for Box {box_number}",
                    data=pdf_data,
                    file_name=pdf_filename,
                    mime="application/pdf",
                )

            # 一键下载所有 PDF（zip）
            if pdf_files:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zipf:
                    for fname, fdata in pdf_files:
                        zipf.writestr(fname, fdata)
                zip_buffer.seek(0)
                st.download_button(
                    label="Download all boxes PDF (ZIP)",
                    data=zip_buffer,
                    file_name="all_boxes_pdf.zip",
                    mime="application/zip",
                )

            # 一键下载所有SKU的表格+大号库位PDF
            sku_pdf_files = []
            for row in box_sku_dicts_df.itertuples():
                box_number = getattr(row, "box_number")
                sku = getattr(row, "sku")
                stock_code = getattr(row, "stock_code")
                pdf_filename = f"{box_number}_{sku}_{stock_code}.pdf"
                pdf_data = create_sku_pdf(
                    sku,
                    getattr(row, "quantity"),
                    stock_code,
                    box_number,
                    getattr(row, "box_index"),
                    box_count,
                )
                sku_pdf_files.append((pdf_filename, pdf_data))
            if sku_pdf_files:
                zip_buffer2 = io.BytesIO()
                with zipfile.ZipFile(zip_buffer2, "w") as zipf:
                    for fname, fdata in sku_pdf_files:
                        zipf.writestr(fname, fdata)
                zip_buffer2.seek(0)
                st.download_button(
                    label="Download all SKU stock_code PDF (ZIP)",
                    data=zip_buffer2,
                    file_name="all_sku_stockcode_pdf.zip",
                    mime="application/zip",
                )


if __name__ == "__main__":
    main()
