     _          _ _      __     __   ___           _
 ___| |__   ___/ / |____ \ \   / /  |_ _|_ __   __| | _____  _____ _ __
/ __| '_ \ / _ \ | |_  /  \ \ / /____| || '_ \ / _` |/ _ \ \/ / _ \ '__|
\__ \ | | |  __/ | |/ /    \ V /_____| || | | | (_| |  __/>  <  __/ |
|___/_| |_|\___|_|_/___|    \_/     |___|_| |_|\__,_|\___/_/\_\___|_|


                                                                        (SANS Index Prettifier)
                                                                        
Author: Ananth Gottimukala (she11z)
GitHub: https://github.com/ananth-she11z
LinkedIn: https://www.linkedin.com/in/ananth-she11z

Special thanks to @Anirban Dey and @Scott Hughes

----------------------------------------------------------------------------------------------------------------------------------------------------------------------

Introduction - v11.1

V-Indexer is an offline tool created in Python3 by Ananth Gottimukala (she11z) as an alternative to SANS Voltaire which is a SANS indexing tool
Special thanks to @Anirban Dey and @Scott Hughes

Voltaire is really a good online web-based tool for SANS indexing. But, due to few bugs it's not feasible. V-Indexer tool will give similar output with same columns in .docx and .csv format.

Below are some issues with Voltaire and how I have mitigated in V-Indexer (Added more fantastic features) -

1. Voltaire is web-based tool requires you to login via any account - V-Indexer is an offline python script which can be run on local system and the code is clear and visible

2. Voltaire saves your index data online - V-Indexer don't require to save any data, it just takes an offline index file in .xlsx and process it to give you Index_<sheet name>.docx and Index_<sheet name>.csv index file

3. Voltaire requires you to use first character of every keyword/title in upper case - V-Indexer can digest both lowercase and uppercase (if you mistakenly use both still OK ;)

4. Voltaire requires you to have atlease one keyword for every character (A to Z) - V-Indexer will still process even if you have few characters with no indexing and will mention you about it in .docx file

5. Voltaire fails to index the last character (i.e Z) even if you have indexed - V-Indexer covers every alphabet, number and special character statically feeded without usage of any Regex

6. Voltaire gives multiple output file formats but yet outputs only .docx format - V-Indexer will produce index in two formats ".docx" and ".csv" for user flexibillity

7. Voltair has no options to process specific/multiple sheets - V-Indexer gives you option "-s" where you can either specify any sheet name under your .xlsx file or simply use "all or ALL" with -s option to process all the sheets

8. Voltaire takes only one keyword - V-Indexer gives you freedom to include any number of keywords per row with same description and book/page (-k option lets you mention how many columns you have keywords in your index)

   NOTE: The way V-Indexer works is, if you have a row with two or more keywords having same description and book/page locations, V-Indexer will fetch each keyword    and add it to another new row having the same details.
   In short V-Indexer will process all your multiple keywords in a single column for better visibility. Example below -

   | keyword-1 | keyword-2 | Keyword-N    | Description | Book | Page |
   | container | Docker    | Docker Image | spin me     |  2   | 115  |

   In above example V-Indexer will process it as follows -

   | keyword      | Description | Book | Page |
   | container    | spin me     |  2   | 115  |
   | Docker       | spin me     |  2   | 115  |
   | Docker Image | spin me     | 2    | 115 |

   Now you can find these under both Cc and Dd in your final index files ;)


Requirements -

V-Indexer uses the following modules(please check requirements.txt file)

On your command prompt run this command from same directory: "pip install -r requirements.txt" to install the following modules

xlrd==1.2.0
docx==0.2.4
python-docx==0.8.10
pyfiglet==0.8.post1


Instructions -

1. FORMAT of your .xlsx should be as follows -

    | Keyword-1 | keyword-2 | ............. | keyword-N | Description | Book Number | Page Number |

    NOTE: Keep your header in place. Script will automatically exclude 0th row as header from processing
    NOTE: If you have multiple book/page numbers for one keyword and the output is not desirable with missing "," between page/book numbers then please use "-" as delimiter between page/book numbers (Eg: 115-145) in place of (115,145)
    NOTE: Make sure you have the same above mentioned format in all your sheets

2. V-Indexer will only process ".xlsx" having multiple "sheets" by using -s option (Eg: "-s Sheet1" OR "-s all" OR "-s 'My Sheet'" OR "-s ALL")

3. Output filename will be "Index_<sheet name>.docx" & "Index_<sheet name>.csv" created in the same directory as script and your index file

4. Help menu will look like below - NOTE: If you have spaces in your filename or sheet name, Please use "" to encapsulate the arguments. Example -

     _          _ _      __     __   ___           _
 ___| |__   ___/ / |____ \ \   / /  |_ _|_ __   __| | _____  _____ _ __
/ __| '_ \ / _ \ | |_  /  \ \ / /____| || '_ \ / _` |/ _ \ \/ / _ \ '__|
\__ \ | | |  __/ | |/ /    \ V /_____| || | | | (_| |  __/>  <  __/ |
|___/_| |_|\___|_|_/___|    \_/     |___|_| |_|\__,_|\___/_/\_\___|_|



                                                                        (SANS Index Prettifier)

Author: Ananth Gottimukala (she11z)
GitHub: https://github.com/ananth-she11z
LinkedIn: https://www.linkedin.com/in/ananth-she11z

Special thanks to @Anirban Dey and @Scott Hughes
-------------------------------------------------------------------------------------------------------
Usage: V-Indexer.py [options]

Options:
  --version           show program's version number and exit
  -h, --help          show this help message and exit
  -k KEYWORD_COLUMNS  Enter number of columns you have with keywords in your
                      index
  -f INDEX_FILENAME   Index filename (.xlsx)
  -c COURSE_NAME      Enter which course you are preparing for (Eg: GDAT,
                      GCIH) Default set to "SANS"
  -s SHEET_NAME       Please specify which sheet to process. (Either -s
                      <SheetName> for specific sheet OR -s <all/ALL> to
                      process all available sheets (Make sure all sheets are
                      in the same format with similar column structure)
  -a CASE             Keywords in Upper/Lower/Capitalize case (Eg: -a
                      lower/upper/capitalize)
                      

Contact -

For reporting issues or for any further information, Please contact me at - ananth DOT venk88 AT gmail DOT com





