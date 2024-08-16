# Excel Bar Chart Project

## The Challenge

The dataset was not pivot-friendly. The goal was to create a bar chart with form controls to select a Month and Year range. The challenge was to ensure the chart displays the selected month along with the two preceding and two following months, while staying within the same year and not extending into the previous or next year.

## My Approach

1. **Date Range Creation**: Used form controls to create a selectable date range.
2. **Data Retrieval**: Applied the `OFFSET` function to pull data for the selected month and the two months before and after.
3. **Year Boundary Handling**: Implemented the `IF` function to ensure the range remains within the selected year.
4. **Data Aggregation and Visualization**: Used the `SUMIFS` function to aggregate the data and created the bar chart.

## Summary

This project involved creating a dynamic bar chart in Excel that adjusts to user-selected date ranges and effectively manages year boundaries. It was an engaging exercise in Excel functionality and data visualization.

![Dashboard Screen Shot](https://github.com/user-attachments/assets/f2214556-d1f9-4216-80c1-2028acb7dbb5)
