# Workout_VBA

This workout program is a personal project where I'm opening up for other people to comment or contribute. Given it's developed with Excel/VBA, one needs to download the Excel file "Eric's Workout Tracker_11FEB23.xlsx" and the text file "Workout_Program.txt". On personal devices, you may create your .xlsm, but I choose to separate it on the repository to have all the code on one sheet. 

Below is a function diagram looking at inputs and outputs and how they relate. Some functions are underdeveloped in the code, but important ones are depicted here.

![Function_diagram](https://github.com/jericdw/Workout_VBA/assets/65636464/9d0a1400-bd6f-46af-b386-9ae104882f90)

Naming conventions were personalized and probably need to be refactored to be a serious project. I'm in the process (eventually) of converting this to C# and thinking through object oriented programming techniques to better handle the data and from multiple users. If you look at "Workout_Blazor" that is a template slightly modified to test the SQLite database before I begin doing any of the logic developed in this program. 

Below is the initial dialogue box to input either individual exercise data as part of a daily workout or an entire set of workout data. 

![image](https://github.com/jericdw/Workout_VBA/assets/65636464/a424a65e-9c6f-45d8-b457-c63abf729639)

For individual exercise data input, the exercise coefficients are calculated based on user-weight in the denominator that by design is optimized recursively over time. 

User-input weight is collected every time for individual exercise data.

![image](https://github.com/jericdw/Workout_VBA/assets/65636464/22efe0d4-9608-4c5a-af1b-090b84a35f23)

Following weight input, the program collects the number of exercises historically that have been done in a given workout category. The example below is for "Chest" workouts that were selected on the first dialogue box. Note: If you wish to add a new "Chest" workout, one needs to add the data onto "Sheet2" where this is scraped and collected based on those workout coefficients. This augments every time you add workouts to your recurring exercise-exercise coefficient (Cex) that gets stored and affects future point calculations recursively. Input point formulas are at the bottom of this next image. 

![image](https://github.com/jericdw/Workout_VBA/assets/65636464/c8d04138-a26e-4ee7-a510-fe5c59bdfc1d)

Based on the type of exercise, there are 4 ways (R=Rep, W=Weight, T=Time, D=Distance) ways to measure the coefficient of performance based on how you specify the type of exercise in the collection in "Sheet2". 

More to follow on the recursive nature of the points calculation. The goal is to find an optimum exercise coefficient below 1.000 that stabilizes from workout to workout. This means that consistency is key and the lower the coefficient (when stabilized) the more difficult the exercise reps/weight/time/distance is to maintain. Also, the number of sets is not critical but simply tracked as the sum of the reps is what is entered for each exercise. 

Developing target workout data utilizes some of the more uniquely named functions to predictively tell the user what their exercise coefficients would be if they follow a given plan. Looking at the "Run" data the Distance data is historical until it gets to the red line. then the next 4 distances are predictive in order to get the coefficient to 0.144 as defined in F2 as the Run exercise 'target' for difficulty based on the formula. 

![image](https://github.com/jericdw/Workout_VBA/assets/65636464/f5850313-f655-4d73-b1d3-28162fc9d7d8)

If there are questions about how to use some of the more customized functions, I'm open to renaming but do not have time to go into detail here on the algorithm and how it works. If you have thoughts on how to improve from an object orientation standpoint, I'm interested in developing classes that can handle appropriate level of detail based on the instances that are developed with the data each user creates with each exercise or workout data iteration. 

This project is incomplete/in-transition, but open to others' thoughts and contributions. Thank you!
