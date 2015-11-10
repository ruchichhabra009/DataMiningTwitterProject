import xlrd
import re
from django.utils.encoding import smart_str, smart_unicode
from nltk.tokenize.punkt import PunktSentenceTokenizer
#----------------------------------------------------------------------
def open_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
    # print number of sheets
    print book.nsheets
    # print sheet names
    print book.sheet_names()
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)
    # read a row
    print first_sheet.row_values(0)
    # read a cell
    cell = first_sheet.cell(2,3)
    print cell
    print cell.value
    # read a row slice
    print first_sheet.row_slice(rowx=0,
                                start_colx=0,
                                end_colx=2)
    obama_datasheet = book.sheet_by_name("Obama");
    obama_rows = obama_datasheet.nrows;
    #gives the number of tweets of obama
    print obama_rows;
    for i in range(1,obama_rows):
    	each_column= obama_datasheet.row_values(i);
    	with open("out.txt", 'a') as outfile:
    		#outfile.writelines(smart_str(each_column[3]));
    		tweet = each_column[3];
    		sentiment = each_column[4];
    		#print each_sentiment;
    		if sentiment == 1 or sentiment == -1 or sentiment == 0:

    			processed_tweet = processTweet(tweet);

    #outfile.close();

def processTweet(stringValue):
	
	#tweet_ = PunktSentenceTokenizer().tokenize(stringValue);
	tweet_ = stringValue.lower();
	tweet_ = re.sub('<[A-Za-z\/][^>]*>', ' ', tweet_)
	tweet_ = re.sub('<[A-Za-z\/][^>]*>', ' ', tweet_)
	tweet_ = re.sub(r'pic.twitter.*?$', '', tweet_)
	tweet_ = re.sub(r'pic.twitter.*? ', '', tweet_)               #remove links
	tweet_ = re.sub('((www\.[^\s]+)|(http://[^\s]+))','',tweet_)  #remove url
	tweet_ = re.sub(r'#([^\s])*$', '', tweet_)
	tweet_ = re.sub(r'#([^\s])* ', '', tweet_)
	tweet_ = re.sub(r'\@([^\s])*$', '', tweet_)
	tweet_ = re.sub(r'\@([^\s])* ', '', tweet_)
	tweet_ = re.sub("\d",'',tweet_)                               #remove digits
	tweet_ = re.sub("[!\'.\"%/*$;:\(\):,?]",'',tweet_)            #remove special characters
	tweet_ = re.sub(r'\-',' ',tweet_)                             #replace - with white space
	tweet_ = re.sub(r'\'m', ' am', tweet_)
	tweet_ = re.sub(r'\'d', ' would', tweet_)
	tweet_ = re.sub(r'\'ll', ' will', tweet_)
	tweet_ = re.sub(r'\&', 'and', tweet_)
    #tweet = re.sub(r'\b\w{1,2}\b','', tweet).strip()            #remove 3 or lesser length words
	tweet_ = re.sub('[\s]+', ' ', tweet_)   
	with open("out.txt", 'a') as outfile:
    		outfile.writelines(smart_str(tweet_));


def write_file(path):

    with open(path, 'w') as outfile:
    	outfile.write("Hello world")


#----------------------------------------------------------------------
if __name__ == "__main__":
    path_in = "training-Obama-Romney-tweets.xlsx";
    path_out = "out.txt";
    open_file(path_in);
    #write_file(path_out);