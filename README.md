# Python Project
The project is part of CyberPro Data Analyst Program 2. The key objective of the project is to show the ability to process data by using Numpy and Pandas in Python VS Code to process dataframes and to generate visuals in Matplotlib and Seaborn. The secondary objective (though it would be the primary objective otherwise) is to demonstrate the analytic ability to interpret the results and to summarize these in a presentation with the results pushed to Git Hub.  


## Project Description
This project is based on a Command Line Interface (CLI) and analyzes mobile telecommunications indicators for voice services across regions and income groups in the 137 countries and aggregate regional or income groups across the world for 2008-2023. The database itself is larger but the countries with a lot of data unavailable and years prior to 2008 and 2024 were exclueded from the anlysis deliberately for comparability and to avoid further technical issus with finding and interpreting comparable data. Empty data were not filled-in or extra- / intrapolated because this house considers this fundamentally wrong. 

The project can be re-applied to other dataframes with a few adjustments for dataframes and types of data. Visualization and presentation can be adjusted and expanded as well. I consider presentation skills in the required professional format apparently cannot be completely replaced by experience in programming skills or by AI, while making presentations by experienced analysts manually can produce a lot better results than the result of this project. Convergent knowledge of both some python and presentation skills is what makes value in this case. Also we decided to abstain from filling-in the slides by GPT because that is where human professionals are still way ahead of AI and high-tech. 


## Data Settings 
Data was obtained in *.zip from the public website of International Telecommunications Union, https://datahub.itu.int/  except for summary table with country classification, which wwas downloaded from https://datahub.itu.int/dashboards/idi/?e=ISR&y=2025  

Zip files are saved and unpacked locally because downloading them from the website is inefficient and slower 

Data includes 3 dataframes – mobile voice indicators, population and subscribers and country income and regional classification

The dataframes were merged, data for 2003-2008 was deleted due to data unavailability for some indicators, countries with more that one zero except the last year were dropped. The remaining data for 137 countries is incomplete but is quite representative for the objective  

Data for 2004 is mainly missing, so further analysis shows comparable data analysis for 2008-2023 

New indicators were calculated further in Create_Charts file – 
1. Market Size (ARPU * Subscribers *12) and
2. Penetration Rate (Subscribers / Population) 

Create_Charts file also calculates means and/or totals for country aggregates classified by 
1. Country income 
2. Geographic region and 
3. World aggregate 

and aggregates for the other indicators both initially available and calculated during dataframe pre-processing.    

The resulting dataframe for further manipulations in ITU_Main, ITU_Utilities and Create_Charts files is available in the file ‘ITU_Mobile_Telecoms‘, was called ‘formatted_for_sbrn’ and is saved in *.xlsx

Correlations and regression charts were calculated and plotted in a separate file ITU_Correlations because they have a non-standard layout 

The analysis on the following slides is based only on mobile voice data, for shortcut the missing conclusions are provided by the author without further data support, because the focus of this is python-based data processing capacity rather than full-scale financial statistics analysis 


## How to Run
1. Install dependencies: supporting libraries, which enable the code running are installed in the beginning of each file. If not please reinstall by using pip install <name> or !pip install <name> for Jupiter Notebook 
2. Run the main script: `python ITU_Main.py`and proceed down the menu. To create chart you should first input 0 to load the pre-prosessed dataframe ‘formatted_for_sbrn.xlsx’, then 1 to proceed with charts creation. Other menu options can be run without first loading the dataframe. The menu is intuitive, guides the user through the interface and handles unintended inputs to avoid errors. 
3. Data selection is available safely for 2008-2023, but not for 2024, although the dataframe has part of 2024 indicators for some indicators and some countries. For pie charts a single year has to be selected.  
4. Other supporting files should be opend from the same folder and include:
    - ITU_Utilities, which upload the dataframe, and manages charts, slides and presentations operatins including selecting items to be inlcuded on a chosen slide layout and saving those, selecting slides to be compiled into a presentation, adding slides to an existing presentation deleting slides or presentations. The slides and presentations are prepared in pptx format.  
    - Create_Charts: a function to select data for the chosen key indicators, for the selected years on the selected chart types, saving these in the Charts folder and a tool to select those charts in a preview mode to decide which are good to be included into which types of pptx presentation slides. 


## Key Features
- Dynamic data selection from the dataframe 
- Aggregated insights by region and income group
- Additional key indicators including Market Size and Penetration Rate 
- Dynamic chart creation (line, bar, stacked diagram, 100% stacked diagram, pie, scatter plot)
- Chart preview on a screen for visibility to decise, which slides might be used for which slide 
- Compilation and saving of slides. The Project was designed to handle 4 automated PowerPoint slide layouts generation -
    1. Title/Divider,
    2. Executive Summary/Conclusions,
    3. 2-chart layout and
    4. 3-slide layout.
    That was sufficient for the purposes of the Project. The textboxes were left blank to be fillled in in PowerPoint directly because it is more practical. 
- Compilation and editing of PowerPoint presentations, including compiling slides in to a newly saved pptx presentation or adding slides to an existing presentation.  

## Command Line Interface Menu
Menu:

0. Load dataframe
1. Create new charts
2. Read charts
3. Prepare slides
4. Compile slides into presentations
5. Delete charts and slides
6. Exit

Choose action: 0
✅ DataFrame loaded and formatted.

Choose action: 1

    Available Indicators:
    1. ARPU
    2. Population
    3. Subscribers
    4. Market Size
    5. Penetration Rate
    Select indicator(s) (e.g., '1,3-4'):

    Available Years:
    1. 2008
    ...
    16. 2023
    Select years (e.g., 'all', '2008-2013', '2008,2010,2012'):

    Select chart type:
    1. Line
    2. Bar
    3. Stacked Column
    4. 100% Stacked Column
    5. Pie
    6. Scatter
    Select chart type (1-6 or name):

    Select countries/regions:
    1. Albania
    ...
    137. Zambia
    138. High-income
    139. Low-income
    140. Lower-middle-income
    141. Upper-middle-income
    142. Africa
    143. Americas
    144. Arab States
    145. Asia-Pacific
    146. CIS
    147. Europe
    148. World
    Enter numbers (e.g., 1,3-5):

Choose action: 2 

    Select files from the pop up menu 


Choose action: 3

    Choose slide layout: 
    1. Cover Slide
    2. Executive Summary 
    3. 2-chart Slide
    4. 3-chart Slide

Choose action: 4

    Select an option:
    1. Create a new presentation from selected slides.
    2. Insert slide(s) into an existing presentation.
   
    Enter option number (1 or 2): 1

        Available Slide Files:
        1. slide_layout_1_1.pptx
        ...
        Enter slide numbers in desired order (e.g., 1,2,4-5):

    Enter option number (1 or 2): 2

        Available Presentations:
        1. Presentation1.pptx
        2. Presentation2.pptx
        3. Updated_Presentation3.pptx
        Select a presentation number to insert into: 3

            Available Slide Files:
            1. slide_layout_1_1.pptx
            ...
            Enter slide numbers to insert (e.g., 1,3-4):

            Enter slide numbers to insert (e.g., 1,3-4): 2
                Presentation has 11 slides.
                1. Adjust Title
                2. Adjust Title
                3. Adjust Title
                4. Adjust Title
                5. Adjust Title
                6. Adjust Title
                7. Adjust Title
                8. Adjust Title
                9. Adjust Title
                10. Adjust Title
                11. Adjust Title
                Insert after which slide number(s)? (e.g., 2,4):

            Insert after which slide number(s)? (e.g., 2,4): 3
                ✅ Slides inserted and saved as Updated_Updated_Presentation3.pptx in 'Presentations' folder.


Choose action: 5

    Which folder would you like to manage?
    1. Charts (\Charts)
    2. Slides (\Slides)
    3. Presentations (\Presentations)

    Enter your choice: 3
        Files in Presentations:
        1. Presentation1.pptx
        2. Presentation2.pptx
        3. Updated_Presentation3.pptx
        4. Updated_Updated_Presentation3.pptx

        Presentation management options:
        1. Delete Slides from Presentation
        2. Delete Entire Presentation

        Enter your choice (1 or 2): 1
            Select presentation number: 3

            Slides in Updated_Presentation3.pptx:
            1. Untitled Slide
            2. Untitled Slide
            3. Untitled Slide
            4. Untitled Slide
            5. Untitled Slide
            6. Untitled Slide
            7. Untitled Slide
            8. Untitled Slide
            9. Untitled Slide
            10. Untitled Slide
            11. Untitled Slide
            Enter slide numbers to delete (e.g., 1,3-4):

        Enter your choice (1 or 2): 2
        Select presentation number: 

6. Exit


## Insights
- The golden era of mobile voice market was mainly completed growth till 2012 shifting to data and messengers from 2013 and on 
- Despite the common expectation for the analyzed market phase the growth of subscribers leads to the decline of mobile voice market, however, it means that revenues shift to data and media market segments 
- Part of revenues was also lost to social media and mobile apps driving EBITDA margins from the peaks of 52%-58% (higher than FAANGS) down to 32%-38% range 
- Mobile operators are seeking for new content- and service- based growth models to avoid becoming a data traffic pipeline with quite varying but still limited success (from negative to within 10% EBITDA margin effect)
- Mobile penetration rates exceeded 100% in most countries due to business voice communications, where businesses and employees have additional SIM cards  
- Less developed emerging markets started the shift towards data before the market saturation due to apparent cost savings 
- For the future of mobile communications market data, media and value chain segments have to be taken into the analysis but they were not in the scope of this presentation  


## What I Learned
- Processing Dataframes with numpy, pandas and doing visuals with matplotlib and seaborn
- Creating 6-types of professionally-looking charts (not just Matplotlib) flexibly from a pre-processed dataframe in a matter of a few menu selections 
- Making standardized slides from the charts in an automated way in a few menu selections with `python-pptx`  
- Compiling presentations and editing slides in presentations in a few menu selections. 
- Handling user input via command-line interface (CLI) 
- Wrapping up the project and pushing the results to Git Hub 
Please mind that several features including inserting other graphics, changing the slide order, printing to pdf or filling-in text by GPT were not included as impractical because these are usually either handled in Microsoft Office or added by a professional more qualified than Chat GPT.  


## Unexpected Trends
- Some regions with high ARPU still had declining penetration.
- Despite the common expectation for the analyzed market phase the growth of subscribers leads to the decline of mobile voice market, however, it means that revenues shift to data and media market segments 


> Created as part of the Cyberpro Data Analyst Program.






