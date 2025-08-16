import os
import re
import math
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
from matplotlib.colors import to_rgb, to_hex
import colorsys
import seaborn as sns
from seaborn import light_palette, dark_palette
from adjustText import adjust_text
import tkinter as tk
from tkinter import filedialog, Tk, Toplevel, Label, Button 
from PIL import Image, ImageTk
import itertools



# from ITU_Utilities import CHARTS_PATH, SLIDES_PATH, PRESENTATIONS_PATH 
# local fallback
CHARTS_PATH = os.path.join(os.getcwd(), "Charts")



plt.rcParams.update({
    'font.size': 20,
    'axes.titlesize': 22,
    'axes.labelsize': 20,
    'xtick.labelsize': 18,
    'ytick.labelsize': 18,
    'legend.fontsize': 18,
    'legend.title_fontsize': 20,
    'font.family': 'Calibri'
})

indicator_labels = {
    "ARPU": "ARPU (US$)",
    "Market Size": "Market Size (Billions)",
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

    # --- Select indicator(s)
    print("\nAvailable Indicators:")
    for key, val in indicators.items():
        print(f"{key}. {val}")
    indicator_input = input("Select indicator(s) (e.g., '1,3-4'): ").strip()
    try:
        selected_inds = set()
        for part in indicator_input.split(','):
            if '-' in part:
                start, end = part.split('-')
                selected_inds.update(range(int(start), int(end) + 1))
            else:
                selected_inds.add(int(part.strip()))
        selected_indicators = [indicators[str(i)] for i in selected_inds if str(i) in indicators]
    except Exception as e:
        print(f" Invalid indicator input: {e}")
        return

    # --- Filter years
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
                    selected_years.update(range(int(start), int(end) + 1))
                else:
                    selected_years.add(int(part.strip()))
            selected_years = sorted(set(selected_years).intersection(set(years)))
    except Exception as e:
        print(f" Invalid year selection: {e}")
        return

    # --- Chart type
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
        "4": "100_stacked", "100%": "100_stacked",
        "5": "pie", "pie": "pie",
        "6": "scatter", "scatter": "scatter"
    }
    chart_type = chart_type_map.get(chart_type_input, "line")
    if chart_type == "pie" and len(selected_years) != 1:
        print(" Pie chart requires exactly one year.")
        return

    # --- Country/region selection
    countries = sorted(df["Country"].dropna().unique())
    income_groups = sorted(df["WB Income Group"].dropna().unique())
    regions = sorted(df["ITU Region"].dropna().unique())
    all_options = countries + income_groups + regions + ["World"]

    print("\n Select countries/regions:")
    for idx, name in enumerate(all_options, 1):
        print(f"{idx}. {name}")
    country_input = input("Enter numbers (e.g., 1,3-5): ").strip()
    try:
        selected_nums = set()
        for part in country_input.split(','):
            if '-' in part:
                start, end = part.split('-')
                selected_nums.update(range(int(start), int(end) + 1))
            else:
                selected_nums.add(int(part.strip()))
        selected_names = [all_options[i - 1] for i in selected_nums if 1 <= i <= len(all_options)]
    except Exception as e:
        print(f" Invalid selection: {e}")
        return

    # --- Prepare combined chart data
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
            val["Indicator"] = selected_indicator
            chart_data = pd.concat([chart_data, val], ignore_index=True)

        if chart_data.empty:
            print(f" No data for '{selected_indicator}'.")
            continue
        if selected_indicator in SUM_INDICATORS:
            chart_data["Value"] /= 1_000_000
        if selected_indicator == "Market Size":
            chart_data["Value"] = chart_data["Value"] / 1_000
        if selected_indicator == "Penetration Rate":
            chart_data["Value"] *= 100
        combined_chart_data = pd.concat([combined_chart_data, chart_data], ignore_index=True)

    if combined_chart_data.empty:
        print("❌ No data for any selected indicator.")
        return

    # --- Color palette
    base_colors = ["#C00000", "#FF6600", "#203864"]

    # --- Color palette: proportional lighter shades
    def create_mixed_color_grades_from_base(base_colors, grades_per_color=5, sets=3):
        """
        Start with original base colors, then generate progressively lighter shades.
        """
        all_colors = []

        for s in range(sets):
            for i in range(grades_per_color):
                for color in base_colors:
                    # Convert hex to RGB
                    rgb = tuple(int(color[j:j+2], 16)/255 for j in (1, 3, 5))
                    h, l, s_ = colorsys.rgb_to_hls(*rgb)

                    # Compute proportional lightness increment
                    # Lighter shades gradually go from current lightness to 1.0
                    step_multiplier = 1.8  # >1 = bigger steps, <1 = smaller steps
                    light_factor = l + (((1 - l) / (grades_per_color * sets - 1)) * (i + s*grades_per_color)) * step_multiplier

                    new_l = min(1.0, light_factor)

                    # Convert back to RGB and append
                    new_rgb = colorsys.hls_to_rgb(h, new_l, s_)
                    all_colors.append(to_hex(new_rgb))

        return all_colors


    # --- Usage
    full_colors = create_mixed_color_grades_from_base(base_colors, grades_per_color=6, sets=3)


    combined_chart_data["Hue"] = combined_chart_data["Country"] + " - " + combined_chart_data["Indicator"]
    hue_list = combined_chart_data["Hue"].unique()
    color_repeat = (len(hue_list) // len(full_colors)) + 1
    palette_dict = dict(zip(hue_list, (full_colors * color_repeat)[:len(hue_list)]))

    # --- Y label
    if chart_type == "100_stacked":
        y_label = "Share (%)"
    elif len(selected_indicators) <= 2:
        y_label = indicator_labels.get(selected_indicators[0], "Value")
    else:
        y_label = "Indicator Value"

    plt.figure(figsize=(14, 6))

    # --- Chart plotting
    if chart_type == "line":
        ax = sns.lineplot(data=combined_chart_data, x="Year", y="Value", hue="Hue", marker="D", palette=palette_dict)
    elif chart_type == "bar":
        ax = sns.barplot(data=combined_chart_data, x="Year", y="Value", hue="Hue", palette=palette_dict)
    elif chart_type in ["stacked", "100_stacked"]:
        pivot_df = combined_chart_data.pivot(index="Year", columns="Hue", values="Value").fillna(0)
        if chart_type == "100_stacked":
            pivot_df = pivot_df.div(pivot_df.sum(axis=1), axis=0) * 100
        ax = pivot_df.plot(kind="bar", stacked=True, figsize=(14, 6),
                           color=[palette_dict.get(col, None) for col in pivot_df.columns])
        if chart_type == "100_stacked":
            ax.yaxis.set_major_formatter(mtick.PercentFormatter(xmax=100))
            ax.set_ylim(0, 100)
    elif chart_type == "scatter":
        ax = sns.scatterplot(data=combined_chart_data, x="Year", y="Value", hue="Hue", palette=palette_dict)
    elif chart_type == "pie":
        pie_year = selected_years[0]
        pie_data = combined_chart_data[combined_chart_data["Year"] == pie_year]
        pie_data_grouped = pie_data.groupby("Hue")["Value"].sum().sort_values(ascending=False)
        colors = [palette_dict.get(hue, "#999999") for hue in pie_data_grouped.index]
        total = pie_data_grouped.sum()
        combined_labels = [f"{name} {value/total:.1%}" for name, value in zip(pie_data_grouped.index, pie_data_grouped)]
        fig, ax = plt.subplots(figsize=(14, 6))
        pie_result = ax.pie(pie_data_grouped, labels=combined_labels, colors=colors,
                            startangle=30, pctdistance=1.15, labeldistance=1.25)
        ax.axis('equal')
        plt.tight_layout()

    # --- Y axis formatting
    if chart_type != "pie":
        if chart_type == "100_stacked":
            ax.yaxis.set_major_formatter(mtick.PercentFormatter(xmax=100))
        elif len(selected_indicators) == 1:
            si = selected_indicators[0]
            if si == "Penetration Rate":
                ax.yaxis.set_major_formatter(mtick.PercentFormatter())
            elif si in SUM_INDICATORS:
                ax.yaxis.set_major_formatter(mtick.FuncFormatter(lambda x, _: f'{x:.0f}M'))
            elif si == "ARPU":
                ax.yaxis.set_major_formatter(mtick.FormatStrFormatter("%.1f"))
        ax.set_ylabel(y_label)
        ax.set_xlabel("Year")

    # --- Unified legend placement
    if chart_type != "pie":
        ncol = 3 if len(hue_list) > 3 else len(hue_list)
        ax.legend(title="", frameon=False, loc='upper center',
                  bbox_to_anchor=(0.5, -0.2), ncol=ncol,
                  fontsize='small', handletextpad=0.5,
                  columnspacing=1.0, borderaxespad=0.5)
        plt.subplots_adjust(bottom=0.35)

    plt.xticks(rotation=0)
    plt.tight_layout()

    # --- Save charts
    safe_indicators = '_'.join(re.sub(r'\W+', '', ind) for ind in selected_indicators)
    safe_countries = '_'.join(re.sub(r'\W+', '', c) for c in selected_names)
    filename_base = f"{safe_indicators}_{safe_countries}"
    existing = [f for f in os.listdir(charts_path) if f.startswith(filename_base) and f.endswith(".jpeg")]
    chart_num = len(existing) + 1
    base_filename = f"{filename_base}_{chart_num}"
    jpeg_path = os.path.join(charts_path, f"{base_filename}.jpeg")
    png_path = os.path.join(charts_path, f"{base_filename}.png")
    plt.savefig(jpeg_path, format='jpeg', dpi=300, bbox_inches='tight')
    plt.savefig(png_path, format='png', dpi=300, bbox_inches='tight')
    plt.show()
    print(f"✅ Chart saved as:\n- {jpeg_path}\n- {png_path}")








    





def select_image_files(charts_path=CHARTS_PATH):
    """
    Open a file dialog to select image files from the given charts_path folder.
    Returns a list of selected file paths.
    """
    root = Tk()
    root.withdraw()

    # Force root visible just enough to bring dialog to front
    root.deiconify()
    root.lift()
    root.focus_force()
    root.attributes('-topmost', True)
    root.update()
    root.after_idle(root.attributes, '-topmost', False)

    file_paths = filedialog.askopenfilenames(
        title="Select Image Files for Slide",
        initialdir=charts_path,
        filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp *.gif")]
    )

    root.withdraw()  # hide root again, but keep alive
    return list(file_paths)





def display_images(image_paths):
    """
    Display selected image files in a scrollable Tkinter window.
    """
    if not image_paths:
        print("No images selected.")
        return

    window = Toplevel()
    window.title("Selected Images Preview")
    window.geometry("900x600")  # give a decent default size

    # --- Scrollable Frame Setup ---
    canvas = tk.Canvas(window)
    scrollbar_y = tk.Scrollbar(window, orient="vertical", command=canvas.yview)
    scrollbar_x = tk.Scrollbar(window, orient="horizontal", command=canvas.xview)

    scroll_frame = tk.Frame(canvas)

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar_y.grid(row=0, column=1, sticky="ns")
    scrollbar_x.grid(row=1, column=0, sticky="ew")

    window.grid_rowconfigure(0, weight=1)
    window.grid_columnconfigure(0, weight=1)

    # --- Display Images ---
    photo_images = []
    for idx, img_path in enumerate(image_paths):
        try:
            img = Image.open(img_path)
            img.thumbnail((400, 300))
            photo = ImageTk.PhotoImage(img)
            photo_images.append(photo)

            label = Label(scroll_frame, image=photo, text=os.path.basename(img_path), compound="top")
            label.grid(row=idx // 2, column=idx % 2, padx=10, pady=10)
        except Exception as e:
            print(f"Failed to open image {img_path}: {e}")

    # Prevent garbage collection of images
    window.photo_images = photo_images

    # Close button
    close_button = Button(scroll_frame, text="Close", command=window.destroy)
    close_button.grid(row=(len(image_paths) + 1) // 2, column=0, columnspan=2, pady=10)



    # Bring to front
    window.lift()
    window.focus_force()
    window.attributes('-topmost', True)
    window.after_idle(window.attributes, '-topmost', False)

    # Wait until closed, then return to CLI
    window.wait_window()



def read_entry(image_paths=None):
    """
    Wrapper function:
    - If image_paths provided, display them.
    - If not, open file dialog and then display.
    """
    if not image_paths:
        image_paths = select_image_files()
    display_images(image_paths)








# === Module Guard ===
if __name__ == "__main__":
    print("This is a helper module. Please run ITU_Main.py instead.")








