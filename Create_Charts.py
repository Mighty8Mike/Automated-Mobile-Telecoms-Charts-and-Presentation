import os
import re
import math
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
from matplotlib.colors import to_hex
import seaborn as sns
from seaborn import light_palette, dark_palette
from adjustText import adjust_text
import tkinter as tk
from tkinter import filedialog, Tk, Toplevel, Label, Button 
from PIL import Image, ImageTk



from ITU_Utilities import CHARTS_PATH, SLIDES_PATH, PRESENTATIONS_PATH 



plt.rcParams.update({
    'font.size': 20,
    'axes.titlesize': 22,
    'axes.labelsize': 20,
    'xtick.labelsize': 18,
    'ytick.labelsize': 18,
    'legend.fontsize': 18,
    'legend.title_fontsize': 20,
    'font.family': 'Calibri'  # or 'Arial', 'DejaVu Sans', etc.
})



indicator_formatting = {
    "ARPU": lambda x: x,
    "Market Size": lambda x: x / 1_000_000,
    "Penetration Rate": lambda x: x / 100,
    "Population": lambda x: x / 1_000_000,
    "Subscribers": lambda x: x / 1_000_000,
}

indicator_labels = {
    "ARPU": "ARPU (US$)",
    "Market Size": "Market Size (Millions)",
    "Penetration Rate": "Penetration Rate (%)",
    "Population": "Population (Millions)",
    "Subscribers": "Subscribers (Millions)",
}






def create(df, charts_path):

    indicators = {
        "1": "ARPU",
        "2": "Population",
        "3": "Subscribers",
        "4": "Market Size",
        "5": "Penetration Rate"
    }

    SUM_INDICATORS = ["Population", "Subscribers", "Market Size"]

    print("\nAvailable Indicators:")
    for key, val in indicators.items():
        print(f"{key}. {val}")
    indicator_input = input("Select indicator(s) (e.g., '1,3-4'): ").strip()

    try:
        selected_inds = set()
        for part in indicator_input.split(','):
            if '-' in part:
                start, end = part.split('-')
                selected_inds.update(range(int(start), int(end)+1))
            else:
                selected_inds.add(int(part.strip()))
        selected_indicators = [indicators[str(i)] for i in selected_inds if str(i) in indicators]
    except Exception as e:
        print(f"❌ Invalid indicator input: {e}")
        return

    df = df.copy()
    df["Year"] = pd.to_numeric(df["Year"], errors='coerce')
    df.dropna(subset=["Year"], inplace=True)
    df["Year"] = df["Year"].astype(int)

    years = sorted(df["Year"].unique())
    print("\nAvailable Years:")
    for i, yr in enumerate(years, 1):
        print(f"{i}. {yr}")
    year_input = input("Select years (e.g., 'all', '2008-2013', '2008,2010,2012'): ").strip()

    try:
        if year_input.lower() == "all":
            selected_years = years
        else:
            selected_years = set()
            for part in year_input.split(','):
                if '-' in part:
                    start, end = part.split('-')
                    selected_years.update(range(int(start), int(end)+1))
                else:
                    selected_years.add(int(part.strip()))
            selected_years = sorted(set(selected_years).intersection(set(years)))
    except Exception as e:
        print(f"❌ Invalid year selection: {e}")
        return

    print("\nSelect chart type:")
    print("1. Line")
    print("2. Bar")
    print("3. Stacked Column")
    print("4. 100% Stacked Column")
    print("5. Pie")
    print("6. Scatter")
    chart_type_input = input("Select chart type (1-6 or name): ").strip().lower()

    chart_type_map = {
        "1": "line", "line": "line",
        "2": "bar", "bar": "bar",
        "3": "stacked", "stacked column": "stacked",
        "4": "100_stacked", "100%": "100_stacked", "100 stacked": "100_stacked", "100% stacked column": "100_stacked",
        "5": "pie", "pie": "pie",
        "6": "scatter", "scatter": "scatter"
    }
    chart_type = chart_type_map.get(chart_type_input, "line")

    if chart_type == "pie" and len(selected_years) != 1:
        print("❌ Pie chart requires exactly one year.")
        return

    countries = sorted(df["Country"].dropna().unique())
    income_groups = sorted(df["WB Income Group"].dropna().unique())
    regions = sorted(df["ITU Region"].dropna().unique())
    all_options = countries + income_groups + regions + ["World"]

    base_colors = ["#C00000", "#FF6600", "#203864"]
    shades = (
        light_palette("#C00000", 5, reverse=True)[1:] +
        dark_palette("#FF6600", 5, reverse=True)[1:] +
        light_palette("#203864", 5, reverse=True)[1:]
    )
    full_colors = base_colors + [to_hex(c) for c in shades]

    print("\n➡️ Select countries/regions:")
    for idx, name in enumerate(all_options, 1):
        print(f"{idx}. {name}")
    country_input = input("Enter numbers (e.g., 1,3-5): ").strip()

    try:
        selected_nums = set()
        for part in country_input.split(','):
            if '-' in part:
                start, end = part.split('-')
                selected_nums.update(range(int(start), int(end)+1))
            else:
                selected_nums.add(int(part.strip()))
        selected_names = [all_options[i-1] for i in selected_nums if 1 <= i <= len(all_options)]
    except Exception as e:
        print(f"❌ Invalid selection: {e}")
        return

    # Prepare combined chart data for all indicators
    combined_chart_data = pd.DataFrame()

    for selected_indicator in selected_indicators:
        df_filtered = df[df["Year"].isin(selected_years)]
        chart_data = pd.DataFrame()

        for name in selected_names:
            if name in countries:
                sub = df_filtered[df_filtered["Country"] == name]
            elif name in income_groups:
                sub = df_filtered[df_filtered["WB Income Group"] == name]
            elif name in regions:
                sub = df_filtered[df_filtered["ITU Region"] == name]
            elif name == "World":
                sub = df_filtered
            else:
                continue

            if selected_indicator == "ARPU":
                ms = sub[sub["Key Indicator"] == "Market Size"].groupby("Year")["Value"].sum()
                subs = sub[sub["Key Indicator"] == "Subscribers"].groupby("Year")["Value"].sum()
                val = (ms / subs / 12).reset_index(name="Value")
            elif selected_indicator == "Penetration Rate":
                subs = sub[sub["Key Indicator"] == "Subscribers"].groupby("Year")["Value"].sum()
                pop = sub[sub["Key Indicator"] == "Population"].groupby("Year")["Value"].sum()
                val = (subs / pop).reset_index(name="Value")
            else:
                sub = sub[sub["Key Indicator"] == selected_indicator]
                if selected_indicator in SUM_INDICATORS:
                    val = sub.groupby("Year")["Value"].sum().reset_index()
                else:
                    val = sub.groupby("Year")["Value"].mean().reset_index()

            val["Country"] = name
            val["Indicator"] = selected_indicator  # add indicator label
            chart_data = pd.concat([chart_data, val], ignore_index=True)

        if chart_data.empty:
            print(f"❌ No data found for '{selected_indicator}'.")
            continue

        if selected_indicator in SUM_INDICATORS:
            chart_data["Value"] /= 1_000_000
        if selected_indicator == "Penetration Rate":
            chart_data["Value"] *= 100

        combined_chart_data = pd.concat([combined_chart_data, chart_data], ignore_index=True)

    if combined_chart_data.empty:
        print("❌ No data found for any selected indicator.")
        return

    # Combine country and indicator for hue
    combined_chart_data["Hue"] = combined_chart_data["Country"] + " - " + combined_chart_data["Indicator"]
    hue_list = combined_chart_data["Hue"].unique()

    color_repeat = (len(hue_list) // len(full_colors)) + 1
    palette_dict = dict(zip(hue_list, (full_colors * color_repeat)[:len(hue_list)]))

    # Set Y-axis label based on selected indicators
    if chart_type == "100_stacked":
        y_label = "Share (%)"
    elif len(selected_indicators) <= 2:
        single_indicator = selected_indicators[0]
        y_label = indicator_labels.get(single_indicator, "Value")
    else:
        y_label = "Indicator Value"

    plt.figure(figsize=(14, 6))

    # Chart rendering based on type
    if chart_type == "line":
        ax = sns.lineplot(data=combined_chart_data, x="Year", y="Value", hue="Hue", marker="D", palette=palette_dict)
    elif chart_type == "bar":
        ax = sns.barplot(data=combined_chart_data, x="Year", y="Value", hue="Hue", palette=palette_dict)
    elif chart_type in ["stacked", "100_stacked"]:
        if len(selected_indicators) > 1:
            print("⚠️ Stacked/100% stacked chart type only supports a single indicator. Using line chart instead.")
            ax = sns.lineplot(data=combined_chart_data, x="Year", y="Value", hue="Hue", marker="D", palette=palette_dict)
        else:
            pivot_df = combined_chart_data.pivot(index="Year", columns="Hue", values="Value").fillna(0)
            pivot_df = pivot_df.sort_index()

            if chart_type == "100_stacked":
                pivot_df = pivot_df.div(pivot_df.sum(axis=1), axis=0) * 100  # Convert to percentage

            ax = pivot_df.plot(
                kind="bar",
                stacked=True,
                figsize=(14, 6),
                color=[palette_dict.get(col, None) for col in pivot_df.columns]
            )

            if chart_type == "100_stacked":
                ax.yaxis.set_major_formatter(mtick.PercentFormatter(xmax=100))
                ax.set_ylim(0, 100)
                y_label = "Share (%)"
            else:
                y_label = indicator_labels.get(selected_indicators[0], "Value")

            ax.set_ylabel(y_label)
            ax.set_xlabel("Year")

            # 🎯 Move legend below the chart with multiple rows if needed
            n_items = len(pivot_df.columns)
            ncol = 4 if n_items > 4 else n_items  # Auto adjust to number of columns

            ax.legend(
                title="",
                frameon=False,
                loc="upper center",
                bbox_to_anchor=(0.5, -0.25),
                ncol=ncol,
                fontsize="small",
                handletextpad=0.5,
                columnspacing=1.0,
                borderaxespad=0.5
            )

            plt.xticks(rotation=0)
            plt.subplots_adjust(top=0.88, bottom=0.30)  # Adjust spacing for large legends



    elif chart_type == "scatter":
        ax = sns.scatterplot(data=combined_chart_data, x="Year", y="Value", hue="Hue", palette=palette_dict)
    elif chart_type == "pie":
        if len(selected_indicators) != 1:
            print("❌ Pie chart supports only one indicator at a time.")
            return

        pie_year = selected_years[0]
        pie_data = combined_chart_data[combined_chart_data["Year"] == pie_year]

        if pie_data.empty:
            print("❌ No data available for selected year and indicator.")
            return

        pie_data_grouped = pie_data.groupby("Hue")["Value"].sum().sort_values(ascending=False)
        total = pie_data_grouped.sum()
        percentages = (pie_data_grouped / total) * 100

        colors = [palette_dict.get(hue, "#999999") for hue in pie_data_grouped.index]

        fig, ax = plt.subplots(figsize=(10, 10))
        wedges, _ = ax.pie(
            pie_data_grouped,
            startangle=90,
            colors=colors,
            wedgeprops=dict(width=0.5)  # Optional donut look
        )

        texts = []

        for wedge, label, pct in zip(wedges, pie_data_grouped.index, percentages):
            angle = (wedge.theta2 + wedge.theta1) / 2
            x = math.cos(math.radians(angle))
            y = math.sin(math.radians(angle))

            ha = 'left' if x > 0 else 'right'
            text = ax.text(
                x * 1.3, y * 1.3,
                f"{label}\n{pct:.1f}%",
                ha=ha, va='center',
                fontsize=13
            )

            # Arrow from text to wedge
            ax.annotate(
                '', xy=(x * 0.7, y * 0.7), xytext=(x * 1.25, y * 1.25),
                arrowprops=dict(arrowstyle="-", color='gray', lw=1.0,
                                connectionstyle="angle3,angleA=0,angleB=90")
            )

            texts.append(text)

        # Auto adjust text to avoid overlap
        adjust_text(texts, expand_text=(1.05, 1.2), arrowprops=dict(arrowstyle="-", color='gray'))

        ax.set_title(f"{selected_indicators[0]} in {pie_year}", fontsize=22)
        ax.axis('equal')
        ax.set_xlabel("")
        ax.set_ylabel("")
        plt.tight_layout()




    # Additional Y-axis formatting for single indicator
    if chart_type == "100_stacked":
        ax.yaxis.set_major_formatter(mtick.PercentFormatter(xmax=100))
        ax.set_ylim(0, 100)
    elif len(selected_indicators) == 1:
        si = selected_indicators[0]
        if si == "Penetration Rate":
            ax.yaxis.set_major_formatter(mtick.PercentFormatter())
        elif si in SUM_INDICATORS:
            ax.yaxis.set_major_formatter(mtick.FuncFormatter(lambda x, _: f'{x:.0f}M'))
        elif si == "ARPU":
            ax.yaxis.set_major_formatter(mtick.FormatStrFormatter("%.1f"))


    # General formatting
    ax.grid(False, axis='x')

    # Calculate optimal ncol for wrapping the legend in two lines
    ncol = math.ceil(len(hue_list) / 2)

    # Place legend at bottom, not top
    ax.legend(
        title="", frameon=False, loc='lower center',
        bbox_to_anchor=(0.5, -0.35), ncol=ncol,
        fontsize='small', handletextpad=0.5,
        columnspacing=1.0, borderaxespad=0.5
    )
    plt.subplots_adjust(bottom=0.35)  # Leave space for legend at bottom
  

    plt.xticks(rotation=0)
    plt.tight_layout()


    # Build filename from indicators and countries/regions
    safe_indicators = '_'.join(re.sub(r'\W+', '', ind) for ind in selected_indicators)
    safe_countries = '_'.join(re.sub(r'\W+', '', c) for c in selected_names)
    filename_base = f"{safe_indicators}_{safe_countries}"

    # Ensure unique filename by checking existing files
    existing = [f for f in os.listdir(charts_path) if f.startswith(filename_base) and f.endswith(".jpeg")]
    chart_num = len(existing) + 1
    base_filename = f"{filename_base}_{chart_num}"

    jpeg_path = os.path.join(charts_path, f"{base_filename}.jpeg")
    png_path = os.path.join(charts_path, f"{base_filename}.png")

    plt.savefig(jpeg_path, format='jpeg', dpi=300, bbox_inches='tight')
    plt.savefig(png_path, format='png', dpi=300, bbox_inches='tight')

    plt.show()
    print(f"✅ Chart saved as:\n- {jpeg_path}\n- {png_path}")









def select_image_files():
    """
    Open a file dialog to select image files from the CHARTS_PATH folder.
    Returns a list of selected file paths.
    """
    root = Tk()
    root.withdraw()
    initial_dir = os.path.join(os.getcwd(), 'Charts')
    file_paths = filedialog.askopenfilenames(
        title="Select Image Files for Slide",
        initialdir=initial_dir,
        filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp *.gif")]
    )
    return list(file_paths)

def display_images(image_paths):
    """
    Display selected image files in a pop-up Tkinter window.
    """
    if not image_paths:
        print("No images selected.")
        return

    window = Tk()
    window.title("Selected Images Preview")

    photo_images = []
    for idx, img_path in enumerate(image_paths):
        try:
            img = Image.open(img_path)
            img.thumbnail((400, 300))
            photo = ImageTk.PhotoImage(img)
            photo_images.append(photo)

            label = Label(window, image=photo)
            label.grid(row=idx // 2, column=idx % 2, padx=10, pady=10)
        except Exception as e:
            print(f"Failed to open image {img_path}: {e}")

    # Prevent garbage collection of images
    window.photo_images = photo_images

    close_button = Button(window, text="Close", command=window.destroy)
    close_button.grid(row=(len(image_paths) + 1) // 2, column=0, columnspan=2, pady=10)

    window.mainloop()

def read_entry(image_paths):
    """
    Wrapper function to call display_images.
    """
    display_images(image_paths)








# === Module Guard ===
if __name__ == "__main__":
    print("This is a helper module. Please run ITU_Main.py instead.")








