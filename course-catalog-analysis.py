""" 
Created on Sun Nov 26 12:54:22 2017
@author: tom

Performs a keyword-based analysis of curriculum data from studiegids
Built for TU Delft to analyse the sustainability content of the education
With minor modifications, will work for any data with similar format

Outputs:
A ranking of all courses, based on the frequency of relevant keywords in the 
free-text fields. Saved to a .txt file in the results directory

Instructions:
todo
"""

# import modules
# general
import os
import logging
import subprocess

# project specific
import nltk

""" Run these 2 lines to download the necessary stopwords data from this library
import nltk
nltk.download('stopwords')
"""


# Order of function definitions
# support/helper
# data import
# semantic analysis
# main body execution code and saving results defined near end


def file_folder_specs(root, uni):
    """ Get file and folder structure - the place to change
    folder name and structure information.
    Returns
    -------
    dict 
        File and folder specs
    """

    files_folders = {
        'root': root,
        'unidata': os.path.abspath(
            os.path.join(root, 'data', uni)),
        'keyworddata': os.path.abspath(
            os.path.join(root, 'data', 'keywords')),
        'results': os.path.abspath(
            os.path.join(root, 'results', uni))
    }

    # todo: check for existence of folders to prevent later errors
    # if not os.path.exists(files_folders[data]): os.makedirs(files_folders[data])
    return files_folders


def _start_logger(logfile='log.txt', filemode='w', detail=False):
    log = logging.getLogger()
    log.setLevel(logging.DEBUG)

    loghandler = logging.FileHandler(logfile, filemode)
    loghandler.setLevel(logging.DEBUG)

    if detail:
        timeformat = logging.Formatter("%(asctime)s %(msecs)d - %(levelname)s - %(module)s.%(funcName)s(%(lineno)d) -"
                                       " %(message)s [%(processName)s(%(process)d) %(threadName)s(%(thread)d)]",
                                       datefmt='%Y%m%d %H%M%S')
    else:
        timeformat = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s", datefmt='%Y%m%d %H%M%S')
    loghandler.setFormatter(timeformat)
    log.addHandler(loghandler)

    return loghandler


def _stop_logger(handler):
    handler.flush()
    handler.close()
    log = logging.getLogger()
    log.removeHandler(handler)


def import_study_gids(ff, course_catalog_fn):
    import pyexcel

    # from data directory, import course_catalog name

    # filepath assumes data stored as data/uni/filename
    filepath = ff['unidata'] + '\\' + course_catalog_fn

    # Get an array from the spreadsheet-based data
    course_catalog = pyexcel.get_array(file_name=filepath, encoding='utf-8')

    headers = course_catalog.pop(0)
    return course_catalog, headers


def import_keywords(ff, keywords_fn):
    filepath = ff['keyworddata'] + '\\' + keywords_fn
    fhand = open(filepath, 'r')
    keywords = [line.strip().lower() for line in fhand]
    fhand.close()
    return keywords


def convert_common_course_names():
    # todo
    """ Some courses have their name changed from year to year
    But it can be useful during analysis/viz to treat these as if they were the same course
    The courses to be renamed could be defined in a txt document
    This function imports and translates the courses to their alternative/common names
    To save and report in the results
    """
    example_translations = dict()
    # best data format for quick addition of conversions?
    # ..probably external text file with 2 columns separated by a delimiter
    example_translations["Advanced Course on LCA"] = "LCA"
    example_translations["Advanced Course on LCA: Theory to Practice"] = "LCA"
    example_translations["LCA Practice & Reporting"] = "LCA"

    return example_translations


def clean_text(text, stopwords):
    """ Converts free-text with punctuation, numbers, capital letters etc.
        into a list of words
        without any of the punctuation and other 'noise'
        and excludes the pre-determined 'stopwords'
    """
    # remove digits
    text = ''.join(i for i in text if not i.isdigit())

    # remove extra whitespace (with split and join), lower case
    text = ' '.join(text.lower().split())

    # forward slash '/' often used to mean 'or'
    text = text.replace('/', ' ')

    # remove punctuation
    import string
    text = text.translate(text.maketrans('', '', string.punctuation))
    # print('\n**after trans**: ', text)

    # remove stop words
    words = [word for word in text.split() if word not in stopwords]
    # print('\n**after splitting into list: ', words)
    return words


def stem_words(words):
    import snowballstemmer as ss
    stemmer = ss.stemmer('english')
    # stemmer = ss.stemmer('dutch')
    word_stems = [stemmer.stemWord(word) for word in words]
    return word_stems


def get_word_frequency(words):
    # blacklisting approach to count all words, not keywords
    counts = dict()
    for word in words:
        # add word with value = 1 or increment word value
        counts[word] = counts.get(word, 0) + 1
    return counts


def get_keyword_frequency(words, keywords):
    # whitelisting approach to only find pre-defined keywords
    counts = dict()
    # print('\n***\nWords = ',  len(text), '\n', text)
    for word in words:
        if word in keywords:
            # print(word)
            counts[word] = counts.get(word, 0) + 1
    return counts


def calculate_metrics(keyword_frequency, word_frequency):
    # to do: extend to include a score for the relevance of each keyword
    # so a broad word like 'energy' can score less than a more specifc one such as 'renewable'
    # this would also require a change to the way keywords are entered and imported
    # to do: multiple keyword-set comparisons? Although not difficult to..
    # run twice on different keyword files and compare the results
    word_metrics = {}
    for course_code, histogram in keyword_frequency.items():
        # print(course_code, word_frequency[course_code])
        word_count = sum(word_frequency[course_code].values())
        keyword_count = sum(histogram.values())
        unqiue_keyword_count = len(histogram)
        # courses with empty text to be filtered out before this point, but check
        if word_count != 0:
            keyword_ratio = keyword_count / word_count
        else:
            keyword_ratio = 0
        #            print("course found with no text(?):", course_code, "\n")
        word_metrics[course_code] = (word_count, keyword_count, keyword_ratio, unqiue_keyword_count)

    return word_metrics


def main():
    # The program code starts here when run
    """ Enter data into config file
    Modes: 
        Methodology: 
            Pre-defined Keyword scoring, keyword-identification
            Time-series or single year
            
        Data:
            If Keyword scoring: input keyword filename
            Course Filters (to include): 
                    Faculties, Programmes, MSc/BSc, Core/Elective,  
            Time Filters (to include):  
                If single year (affects results) 1. Else, list of years
        
        Results:
            Top N words
        
            Visualisations
                Check AIDA booklet
                "VosViewer is particularly good at producing textual maps of any sorts
                not just from scientometric datasets."
                Messaged AIDA employees at 3me via contact form
        
    """

    # SETTINGS
    root = 'C:/code/course-catalog/'
    uni = "delft"
    ff = file_folder_specs(root, uni)
    course_catalog_fn = 'studiegids_1718.xls'
    # course_catalog_fn = 'studiegids.xls'
    keywords_fn = 'sustainability_keywords.txt'
    # keywords_fn = 'waste_keywords.txt'
    # keywords_fn = 'keywords_green_branding.txt'

    # for print to screen
    words_to_show = 5

    # later extention of code - allow comparison between different sets of
    # keywords to understand relative importance of topics across courses
    # e.g. for comparison of social, economic and environmental content prevalence
    comparativeananalysis = False

    # READ DATA

    # words to exclude from analysis

    # stopwords = get_stopwords('dutch')
    stopwords = nltk.corpus.stopwords.words('english')
    # some words were missing from the list which occur frequently in studiegids
    # these should also be removed as they distort the keyword analysis
    custom_stopwords = ['will', 'refer', 'part', 'description',
                        'see', 'can', 'course', 'students', 'assignment', 'o',
                        'us', 'also', 'lecture', 'main', 'module', 'exam',
                        'work', 'week', 'brightspace', 'blackboard']  # 'x000d'
    stopwords += custom_stopwords

    # read data from xlsx file
    course_catalog, headers = import_study_gids(ff, course_catalog_fn)
    print("Courses loaded:", len(course_catalog), "\nTotal headers/columns:", len(headers))

    # import keywords to seek from file
    keywords = import_keywords(ff, keywords_fn)
    print("Keywords imported:", len(keywords))

    # ORGANIZE/CLEAN DATA

    # stem keywords, for later comparison to stemmed free-text words
    # list(set()) removes duplicates
    keyword_stems = list(set(stem_words(keywords)))
    print("Distinct keywords after stemming:", len(keyword_stems))

    # set filter conditions
    # if empty, the filter will effectively be ignored, see 'idx_' lines below

    # ignore courses with 0 ECTS assigned: 41 of 2633 courses for 17/18
    ects_min = 1

    all_faculties = ['', 'CiTG', 'LR', 'TNW', 'BK', 'Extern', '3mE', 'EWI', 'TBM', 'IO', 'UD']
    faculty_include = []  # ['3mE']
    program_include = []
    """studiegids has 2 language settings - we are using the export of the English version
    some of the courses taught in Dutch(Nederlands) only have English info still
    but the data for these is poorly filled in and hence excluded from our language list
    Some small code changes are required for this software to work with other languages:
        edit the stop words, custom stop words, and stemming algorithm, then check results """
    language_include = ['English', 'Engels', 'Engels, Nederlands', 'Nederlands (op verzoek Engels)']  # , 'Nederlands']

    # for each filter, determine the indices of courses which pass that filter
    # if the filter conditions are left empty, all courses pass the filter due to "or" condition
    idx_ects_min = [i for i, item in enumerate(course_catalog) if item[headers.index('ECTS_POINTS')] >= ects_min]
    idx_faculty_include = [i for i, item in enumerate(course_catalog) if
                           item[headers.index('BUREAU_ID')] in faculty_include or faculty_include == []]
    idx_program_include = [i for i, item in enumerate(course_catalog) if
                           item[headers.index('EDUCATION_CODE')] in program_include or program_include == []]
    idx_language_include = [i for i, item in enumerate(course_catalog) if
                            item[headers.index('COURSELANGUAGE')] in language_include or language_include == []]

    # combine all filters using sets, to leave only courses which pass all filters
    idx = list(set(idx_ects_min) & set(idx_faculty_include) & set(idx_program_include) & set(
        idx_language_include))  # & set(idx_contact_exclude))

    # apply filter to select courses from imported catalog
    courses_to_assess = [course_catalog[i] for i in idx]
    print("Courses that pass all filters:", len(courses_to_assess))
    print("Course keyword analysis starting, this may take a few minutes...")

    # the data for Delft was received with the following headers containing free text
    free_text_headers = ['SUMMARY', 'COURSECONTENS', 'COURSECONTENSMORE', 'STUDYGOALS',
                         'STUDYGOALSMORE', 'EDUCATIONMETHOD', 'LITRATURE', 'PRACTICALGUIDE',
                         'BOOKS', 'READER', 'ASSESMENT', 'SPECIALINFORMATION', 'REMARKS']
    # get indices of these headers for later use
    wanted_header_indices = [headers.index(wanted_header) for wanted_header in free_text_headers]

    # empty objects for use within loop
    courses_no_words = []
    word_frequency = {}
    keyword_frequency = {}
    all_clean_words = []
    all_word_stems = []
    course_metadata = {}

    for i, course in enumerate(courses_to_assess):
        free_text = ''
        clean_words = ''
        course_id = course[headers.index('COURSE_ID')]
        # course_code only works as unique identified when looking at a single year
        # So include year with code, as a tuple, separated only at display of results
        # course_code = (course[headers.index('YEAR_LABEL')], course[headers.index('COURSE_CODE')])

        for j in wanted_header_indices:
            # construct a long string with all free text from chosen columns
            free_text = ' '.join([course[int(j)] for j in wanted_header_indices])

        if free_text == '':
            # stop analysis for any courses that have no words that can be analysed
            courses_no_words.append(course_id)
            continue
        else:
            clean_words = clean_text(free_text, stopwords)
            word_stems = stem_words(clean_words)
            # word_frequency takes no account of the keywords, but instead counts all
            word_frequency[course_id] = get_word_frequency(word_stems)
            keyword_frequency[course_id] = get_keyword_frequency(word_stems, keyword_stems)
            # store metadata on relevant courses for easy later access

            course_metadata[course_id] = (
                course[headers.index('COURSE_CODE')].strip(), course[headers.index('YEAR_LABEL')],
                course[headers.index('COURSE_TITLE')].strip(), course[headers.index('EDUCATION_CODE')],
                course[headers.index('BUREAU_ID')]
            )

            # for analysis at a higher level than specific courses.
            all_clean_words += clean_words
            all_word_stems += word_stems

    # to view all original cleaned words and their stemmed forms
    # [print(word, stem) for word, stem in sorted(zip(all_clean_words, all_word_stems))]
    unique_word_stems = set(all_word_stems)
    print("Courses with no free text found:", len(courses_no_words), ". Therefore", len(keyword_frequency),
          "courses for keyword analysis.")
    print("These courses contained", len(all_clean_words), "total words, including", len(unique_word_stems),
          "unique word stems.")

    # get metrics on the words and keywords, for ranking and display
    word_metrics = calculate_metrics(keyword_frequency, word_frequency)

    # FORMAT RESULTS    
    # order the list from highest scoring to least, for display
    word_metrics = sorted(word_metrics.items(), key=lambda x: x[1][2], reverse=True)

    # print and save metrics and common keywords for each course in a table structure
    import datetime
    datestring = datetime.datetime.now().strftime("%Y-%m-%d_%H%M")
    results_fn = datestring + '.txt'  # +faculty_include[0]?
    ff = file_folder_specs(root, uni)
    results_file = ff['results'] + '\\' + results_fn
    fout = open(results_file, 'w')

    header = "CourseCode\tYear\tCourseTitle\tProgramCode\tFaculty\tUniqueKeywords\tKeywords\tWords\tPercent"
    fout.write(header)

    # print("\nShowing most common", words_to_show, "keyword stems each for courses with over",
    #       keyword_percent_threshold, "% keyword prevalence.")
    # print(40 * "--")
    # print(header)
    # print(40 * "--")
    for course_id, stats in word_metrics:
        histogram = keyword_frequency[course_id]
        keyword_percent = round(stats[2] * 100, 1)
        # only show courses that have a sufficient number of keywords
        word_count = stats[0]
        keyword_count = stats[1]
        unqiue_keyword_count = stats[3]
        # frequent_keywords = sorted(histogram, key=histogram.get, reverse=True)[:words_to_show]
        # keywords_string = [k, '\t' for k in frequent_keywords]
        course_code = str(course_metadata[course_id][0])
        course_year = str(course_metadata[course_id][1])
        course_title = str(course_metadata[course_id][2])
        program_code = str(course_metadata[course_id][3])
        faculty = str(course_metadata[course_id][4])
        line = course_code + '\t' + course_year + '\t' + course_title + '\t' + program_code + '\t' + faculty + '\t' + \
            str(unqiue_keyword_count) + '\t' + str(keyword_count) + '\t' + str(word_count) + '\t' + \
            str(keyword_percent)  # + '\t' + str(frequent_keywords)
        fout.write('\n' + line)

    fout.close()

    keywords_not_in_dataset = set(keyword_stems) - set(all_word_stems)
    print("Keywords not found anywhere in dataset:", keywords_not_in_dataset)

    """
    print('\nMost frequent of all words:')
        for course_code, histogram in word_frequency.items():
        total_words = word_count[course_code]
        frequent_words = sorted(histogram, key=histogram.get, reverse=True)[:words_to_show]
        print(course_code, ':', frequent_words, ', total words =', total_words)
    """
    # VISUALIZE

    return locals()


if __name__ == "__main__":
    # The main routine gets only started if the script is run directly. 
    # It only includes the logging boilerplate and a top level try-except for catching and logging all exceptions.

    # START LOGGING
    if not os.path.exists('./log'):
        os.makedirs('./log')
    log_summary = _start_logger(logfile='./log/process.log')
    # log_detail = _start_logger(logfile = './log/process_details.log', detail = True)
    logging.info('Start logging of {}'.format(__file__))

    try:
        logging.info("Current git commit: %s",
                     subprocess.check_output(["git", "log", "--pretty=format:%H", "-n1"]).decode("utf-8"))
    except:
        logging.warning('Running without version control')

    # MAIN PROGRAM
    try:
        # The following update your local namespace with the variables from main()
        locals().update(main())
        # if you don't want the script to pollute your namespace use 
        # results = main()
        # which gives you all varibales from main in a dict called 'results'

    except Exception as exc:
        logging.exception(exc)
        raise
    finally:
        # STOP LOGGER - clean
        _stop_logger(log_summary)
        # _stop_logger(log_detail)
