# LEDSafari - CO2Calculator & Recommender 

Please note : To run, you need a source file in xlsx format. Since we use customers data, we are not keeping it open in public.

## Pre-requisites
Python-3.6
If Python-3.6 is not there, please download [here.](https://www.python.org/downloads/)

### Installation setup
1. Open terminal in Mac or Command prompt in windows.
2. Please create a working directory.
 
 	``` mkdir Recommendations ```
3. Navigation to the Recommendations folder.
	
	```cd Recommendations```

4. Please create a directory or folder called Data.

	``` mkdir Data ```

5. Please place the user file (BetaCalculator_*.xlsx) renamed as 'source.xlsx' under Recommendations/Data folder. (As usual file placement outside Mac terminal)

6. Navigate to the root folder.

	``` cd ..```
7. Download Recommendation-process.py (Clone or Download option in right top) in the Recommendations folder.
8. Please type (For installation of openpyxl module)

	```sudo pip3 install openpyxl``` 
9. Please type

	```python3 Recommendation-process.py```

Output will contain reduced frequency levels in different parameters in the Python window and also in the text file named 'Recommendation-summary.txt'.			
