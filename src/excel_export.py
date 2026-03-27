import pandas as pd


def export_to_excel(df_products, df_customers, df_sales, analysis, company_name, output_file):

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:

        workbook = writer.book

        currency_fmt = workbook.add_format({'num_format': '₦#,##0'})
        percent_fmt = workbook.add_format({'num_format': '0.00%'})

        header_fmt = workbook.add_format({
            'bold': True,
            'bg_color': '#1F4E78',
            'font_color': 'white'
        })

        title_fmt = workbook.add_format({
            'bold': True,
            'font_size': 16,
            'font_color': '#1F4E78'
        })

        na_fmt = workbook.add_format({
            'italic': True,
            'font_color': '#808080'
        })

        header = workbook.add_format({
            'bold': True,
            'font_size': 26,
            'font_color': 'white',
            'align': 'center',
            'bg_color': '#121212'
        })

        kpi_label = workbook.add_format({
            'font_color': 'white',
            'bg_color': '#1F4E78',
            'align': 'center',
            'bold': True
        })

        revenue_value = workbook.add_format({
            'bold': True,
            'font_size': 20,
            'font_color': 'white',
            'bg_color': '#2E75B6',
            'align': 'center',
            'num_format': '₦#,##0'
        })

        integer_value = workbook.add_format({
            'bold': True,
            'font_size': 20,
            'font_color': 'white',
            'bg_color': '#2E75B6',
            'align': 'center'
        })

        cost_profit_value = workbook.add_format({
            'bold': True,
            'font_size': 18,
            'font_color': 'white',
            'bg_color': '#4472C4',
            'align': 'center',
            'num_format': '₦#,##0'
        })

        margin_value = workbook.add_format({
            'bold': True,
            'font_size': 18,
            'font_color': 'white',
            'bg_color': '#5B9BD5',
            'align': 'center',
            'num_format': '0.00%'
        })

        na_kpi_value = workbook.add_format({
            'bold': True,
            'font_size': 18,
            'font_color': '#D9D9D9',
            'bg_color': '#666666',
            'align': 'center'
        })

        # =========================
        # RAW SHEETS
        # =========================
        df_products.to_excel(writer, sheet_name="Products", index=False)
        df_customers.to_excel(writer, sheet_name="Customers", index=False)
        df_sales.drop(columns=["Month_Num"], errors='ignore').to_excel(
            writer, sheet_name="Sales_Data", index=False
        )

        for sheet in ["Products", "Customers", "Sales_Data"]:
            ws = writer.sheets[sheet]
            ws.set_column("A:Z", 18)
            ws.set_row(0, None, header_fmt)

        # =========================
        # ANALYSIS SHEET
        # =========================
        analysis_ws = workbook.add_worksheet("Analysis")
        writer.sheets["Analysis"] = analysis_ws

        analysis_ws.write("A1", "Business Analysis Summary", title_fmt)

        analysis_ws.write("A3", "Total Revenue")
        analysis_ws.write("B3", analysis["total_revenue"], currency_fmt)

        analysis_ws.write("A4", "Total Cost")
        if analysis["total_cost"] is not None:
            analysis_ws.write("B4", analysis["total_cost"], currency_fmt)
        else:
            analysis_ws.write("B4", "N/A", na_fmt)

        analysis_ws.write("A5", "Total Profit")
        if analysis["total_profit"] is not None:
            analysis_ws.write("B5", analysis["total_profit"], currency_fmt)
        else:
            analysis_ws.write("B5", "N/A", na_fmt)

        analysis_ws.write("A6", "Average Margin")
        if analysis["avg_margin"] is not None:
            analysis_ws.write("B6", analysis["avg_margin"] / 100, percent_fmt)
        else:
            analysis_ws.write("B6", "N/A", na_fmt)

        analysis["revenue_by_product"].to_excel(writer, sheet_name="Analysis", startrow=8, index=False)
        analysis["revenue_by_region"].to_excel(writer, sheet_name="Analysis", startrow=23, index=False)
        analysis["monthly_trend"].to_excel(writer, sheet_name="Analysis", startrow=38, index=False)
        analysis["top_customers"].to_excel(writer, sheet_name="Analysis", startrow=53, index=False)

        # Optional profit-by-product section
        profit_section_start = 68
        analysis_ws.write(f"A{profit_section_start}", "Profit by Product", title_fmt)

        if not analysis["profit_by_product"].empty:
            analysis["profit_by_product"].to_excel(
                writer,
                sheet_name="Analysis",
                startrow=profit_section_start,
                startcol=0,
                index=False
            )
        else:
            analysis_ws.write(f"A{profit_section_start + 2}", "N/A", na_fmt)

        analysis_ws.set_column("A:Z", 18)

        # =========================
        # DASHBOARD SHEET
        # =========================
        dashboard = workbook.add_worksheet("Dashboard")
        dashboard.set_column("A:Z", 22)

        dashboard.merge_range("A1:L2", f"{company_name} Business Intelligence Dashboard", header)

        # First row KPIs
        dashboard.write("B5", "Total Revenue", kpi_label)
        dashboard.write("B6", analysis["total_revenue"], revenue_value)

        dashboard.write("D5", "Total Orders", kpi_label)
        dashboard.write("D6", len(df_sales), integer_value)

        dashboard.write("F5", "Total Units Sold", kpi_label)
        dashboard.write("F6", df_sales["Quantity"].sum(), integer_value)

        # Second row KPIs
        dashboard.write("H5", "Total Cost", kpi_label)
        if analysis["total_cost"] is not None:
            dashboard.write("H6", analysis["total_cost"], cost_profit_value)
        else:
            dashboard.write("H6", "N/A", na_kpi_value)

        dashboard.write("J5", "Total Profit", kpi_label)
        if analysis["total_profit"] is not None:
            dashboard.write("J6", analysis["total_profit"], cost_profit_value)
        else:
            dashboard.write("J6", "N/A", na_kpi_value)

        dashboard.write("L5", "Average Margin", kpi_label)
        if analysis["avg_margin"] is not None:
            dashboard.write("L6", analysis["avg_margin"] / 100, margin_value)
        else:
            dashboard.write("L6", "N/A", na_kpi_value)

        # =========================
        # CHART DATA RANGES
        # =========================
        product_rows = len(analysis["revenue_by_product"]) + 9
        region_rows = len(analysis["revenue_by_region"]) + 24
        month_rows = len(analysis["monthly_trend"]) + 39
        customer_rows = len(analysis["top_customers"]) + 54

        # =========================
        # Revenue by Product
        # =========================
        chart1 = workbook.add_chart({'type': 'column'})
        chart1.add_series({
            'categories': f'=Analysis!A10:A{product_rows}',
            'values': f'=Analysis!B10:B{product_rows}',
            'fill': {'color': '#2E75B6'},
            'data_labels': {'value': True}
        })
        chart1.set_title({'name': 'Revenue by Product'})
        chart1.set_y_axis({'major_gridlines': {'visible': False}})
        dashboard.insert_chart("B10", chart1)

        # =========================
        # Revenue by Region
        # =========================
        chart2 = workbook.add_chart({'type': 'column'})
        chart2.add_series({
            'categories': f'=Analysis!A25:A{region_rows}',
            'values': f'=Analysis!B25:B{region_rows}',
            'fill': {'color': '#70AD47'},
            'data_labels': {'value': True}
        })
        chart2.set_title({'name': 'Revenue by Region'})
        chart2.set_y_axis({'major_gridlines': {'visible': False}})
        dashboard.insert_chart("H10", chart2)

        # =========================
        # Monthly Revenue Trend
        # =========================
        chart3 = workbook.add_chart({'type': 'line'})
        chart3.add_series({
            'categories': f'=Analysis!A40:A{month_rows}',
            'values': f'=Analysis!B40:B{month_rows}',
            'line': {'color': '#FFC000'},
            'data_labels': {'value': True}
        })
        chart3.set_title({'name': 'Monthly Revenue Trend'})
        chart3.set_y_axis({'major_gridlines': {'visible': False}})
        dashboard.insert_chart("B25", chart3)

        # =========================
        # Top Customers
        # =========================
        chart4 = workbook.add_chart({'type': 'bar'})
        chart4.add_series({
            'categories': f'=Analysis!A55:A{customer_rows}',
            'values': f'=Analysis!B55:B{customer_rows}',
            'fill': {'color': '#ED7D31'},
            'data_labels': {'value': True}
        })
        chart4.set_title({'name': 'Top Customers'})
        chart4.set_x_axis({'major_gridlines': {'visible': False}})
        dashboard.insert_chart("H25", chart4)

    print(f"\nBI Dashboard Excel generated: {output_file}")