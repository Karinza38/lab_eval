import pandas as pd
import pathlib
import json
import argparse
import configparser

import lab_eval.eval_class
# install odfpy


""" Describes different ways to assign Grades and export them to Genote 

"""

def append_stem( stem_suffix:str, input_path:str, suffix:str=None ):
  """ appends stem_suffix to input_path

    This function is useful when specific file such as score_list
    for example are derived from input_path.
    In our example, score_list.json derived from "./INF808-AK Notes.ods"
    is named "./INF808-AK Notes--score_list.json"
  """
    
  output_path = pathlib.PurePath( input_path )
  output_path = output_path.with_stem( f"{output_path.stem}--{stem_suffix}" )
  if suffix != None:  
    output_path = output_path.with_suffix( suffix )
  return output_path


class GenoteXLSFile:

  """ This class considers the Genote File Format. Genote File are not 
  different from other XLS files but they do have there own parameters 
  that makes them specific.

  """

  def __init__( self, file_path ):
    self.file_path=file_path
    self.student_id_row = 17
    self.student_id_col = 0
    self.export_score_list_name = None
      
  def import_score_list( self, score_list_path:str ):
    score_list_file = lab_eval.eval_class.ScoreList( score_list_path )
    score_list = score_list_file.load_score_list( )  

    ## we need to define a function for apply in the dataframe  
    def get_score( student_id ):
      if student_id in score_list.keys():
        grade = score_list[ student_id ][ "grade (%)" ]
      else:
        grade = 0
      return grade
    ## the engine needs to be specified to read xls (at leats the new format)
    ## we need to read -2 as the first row being read is interpreted as the header
    ## in our case the header contains the name of the column and is overwritten    
    df = pd.read_excel( self.file_path, engine='openpyxl', header=self.student_id_row - 2,\
                        usecols=[ self.student_id_col ], names=[ "students_id" ] )
    df = df.dropna( )
    df[ "grade (%)" ] = df[ "students_id" ].apply(  get_score )
    genote_export_path = append_stem( 'grades', self.file_path, suffix='.xlsx' ) 
    df.to_excel( genote_export_path, engine='openpyxl' )
    self.export_score_list_name = genote_export_path 
    return genote_export_path  


class MoodleXLSReportFile:
  """ Moodle Files exported the Main page of the course under the "Notes (FR)" ( probably "Grade (EN)" link.
      https://moodle.usherbrooke.ca/grade/report/grader/index.php?id=37245
  """
  
  def __init__( self, file_path ):
    self.file_path = file_path
    ## path of the exported score_list file (once created)  
    self.export_score_list_name = None

  def export_score_list( self, note_col:int=6,student_id_col:int=2 ):
    """ Exports notes of note_col into the score_list format

    Args:
      note_col: note column index. When Notes files are exported from Moodle,
        the user can select one or multiple exams. When only a single test has
        been selected, notes are listed in column 6 which is the default value.

        IMPORANT: notes are expected to be expressed as a percentage.
    
      student_id_col: column index with student_id. The second column represents 
        the student_id. In our case the column contains the student_id as well as other information: 
        uid=glya9851,ou=personnes,dc=usherbrooke,dc=ca
    """
    score_list = {}
      
    df = pd.read_excel( pathlib.Path( self.file_path ) )  
    note_list = df.iloc[:, note_col ]
    for i, student_id in enumerate( df.iloc[:, student_id_col ] ) :
     ## only keeping the student_id   
      k = student_id.split( ',' )[ 0 ].split( '=' )[1] 
      n = note_list[ i ] 
      ## score_list contains questions as well as two expressions of the grade
      ## We fill 3 times teh same value to prevent score_list to generate errors when 
      ## further manipulations are performed.  
      score_list[ k ] = { 1 : n,  "grade (total)" : n, "grade (%)" : n }
    score_list_path = append_stem( 'score_list', self.file_path, suffix='.json' ) 
    with open( score_list_path, 'w', encoding='utf8' ) as f:
      f.write( json.dumps( score_list, indent=2 ) )
    self.export_score_list_name = score_list_path
    return score_list_path  

class MoodleJSONExamFile:

  """ This file designates the File exported by Moodle when an Exam is performed.

  In this case, we only considered the exportation with the questions.
  If that does not work, maybe using MoodleReportFile might be considered.
  """

  def __init__( self, file_path ):
    self.file_path = file_path
    ## path of the exported score_list file (once created)
    self.export_score_list_name = None

  def normalized_float( self, float_str:str )-> float:
    """ normalize float writting

    Args:
      v( (str) : the representatiion of the float

    Return:
      the expected float.

    Moodle (or LibreOffice) can write numbers like "2,3" instead of "2.3"
    we ensure the format is appropriated
    In some cases the value is also "-" which is replaced by 0
    """

    if float_str == '-':
      float_str = 0
    else:
      float_str = float_str.replace( ',', '.' )
    return float( float_str )

  def export_score_list( self ):
    """ exporting file into a JSON Score List format

    """
    ## reading the moodle file
    with open( self.file_path, 'rt', encoding='utf8' ) as f:
      moodle_list = json.loads( f.read( ) )[ 0 ]

    score_list = {}
    for moodle_student in moodle_list:
      question_dict = {}
      question_index = 1
      for k in moodle_student.keys() :
        if k[0] == 'q':
          question_dict[ question_index ] = self.normalized_float( moodle_student[ k ] )
          question_index += 1
        elif k == "note10000":
          question_dict[ "grade (%)" ] = self.normalized_float( moodle_student[ 'note10000' ] )
      score_list[ moodle_student[ 'nomdutilisateur' ] ] = question_dict

    score_list_path = append_stem( 'score_list', self.file_path, suffix='.json' )
    with open( score_list_path, 'w', encoding='utf8' ) as f:
      f.write( json.dumps( score_list, indent=2 ) )
    self.export_score_list_name = score_list_path
    return score_list_path

class XLSFile:

  """ XLS File containing Grades associated to Groups. """

  def __init__( self, file_path ):
    self.file_path = file_path

  def XYdf( self, x_col:int, y_col:int, row_header:int, last_row:int, x_name:str, y_name:str ):
    """ return DataFrame of two columns X and Y

      Note that columns are indexed from 0.

      Args:
        x_col (int): index of the column X.
        y_col (int): index of the column Y.
        row_header (int): index of the header row.
        last_row (int): index of the last row.
        x_name (str) : name or label of the 'x' column
        y_name (str) : name or label of the 'y' column
    """

    df = pd.read_excel( self.file_path, header=row_header, usecols=[ x_col, y_col ],
                       nrows=last_row - row_header, names=[ x_name, y_name ] )
    df = df.dropna( )
    return df

  def group_export_score_list( self, moodle_json_group_path, x_col:int, y_col:int, row_header:int, last_row:int ):
    ## collecting the list of student
    with open( moodle_json_group_path, 'rt', encoding='utf8' ) as f:
      student_list = json.loads( f.read() )
    grade_df = self.XYdf( x_col=1, y_col=2, row_header=2, last_row=39, x_name="Group", y_name="Grade" )
    score_list = {}
    for student in student_list[0]:
      student_id = student[ 'nomdutilisateur' ]
      group = student[ 'groupe' ]
      try:
        grade = grade_df.loc[ grade_df[ 'Group' ] == group ][ 'Grade' ].values[ 0 ]
        score_list[ student_id ] = { 1 : grade,  "grade (total)" : grade, "grade (%)" : grade }
      except IndexError:
        print( f"    WARN unable to find Grade or Group ({group}) in XLS file for student {student_id}" )
    score_list_path = append_stem( 'score_list', self.file_path, suffix='.json' )
    with open( score_list_path, 'w', encoding='utf8' ) as f:
      f.write( json.dumps( score_list, indent=2 ) )
    self.export_score_list_name = score_list_path
    return score_list_path




def grade_to_genote():
    
  description = """import grades to Genote from various formats"""

  usage = \
  """
  This section describes how to import the grades to an Genote XLS File
  when the grades are provides in various files. 

  The general usage is:
  grade_to_genote [grade_file options] grade_file genote.xls

  notes-INF808Gr28-A2024.xlsx represents the Genote XLS File

  1) Importing Grades From a JSON Score List: 
  -------------------------------------------

  score_list.json represents the JSON Score List File.

  grade_to_genote --type 'score_list' ./score_list.json  ./notes-INF808Gr28-A2024.xlsx
  grade_to_genote -t 'score_list' ./score_list.json  ./notes-INF808Gr28-A2024.xlsx


  2) Importing Grades From a Moodle XLS Report File (for a single Exam):
  ----------------------------------------------------------------------

  INF808-AK\ Notes.ods represents the Moodle XLS Report File accessed from:
  Notes -> Telecharger (JavaScript Notation)

  grade_to_genote ./INF808-AK\ Notes.ods  ./notes-INF808Gr28-A2024.xlsx
  grade_to_genote --type 'moodle_xls_report' ./INF808-AK\ Notes.ods  ./notes-INF808Gr28-A2024.xlsx
  grade_to_genote --t 'moodle_xls_report' ./INF808-AK\ Notes.ods  ./notes-INF808Gr28-A2024.xlsx


  3) Importing Grades From a Moodle Exam File (With Questions):
  -------------------------------------------------------------

  INF808-AK-Examen\ Intra-notes.json represents the Moodle Exam File 
  (with Questions). It can be accessed from:  Exam -> Resultats -> 
  Télécharger (JavaScript Notation) 

  grade_to_genote.py -t 'moodle_json_exam' INF808-AK-Examen\ Intra-notes.json  ./notes-INF808Gr28-A2024.xlsx
  grade_to_genote.py -type 'moodle_json_exam' INF808-AK-Examen\ Intra-notes.json  ./notes-INF808Gr28-A2024.xlsx

  4) Importing Grades From an XLS File with Group Grades:
  -------------------------------------------------------

  Presentations_lundi.xlsx represents the XLS File were each Group is
  being assigned a Grade. The file containms two columns Group and 
  Grade. The labels (headers) do not matters. 

  INF808-AL_groups.json represents the Moodle JSON Group File. It is 
  accessed from: Participants -> Groupes/Vue d'Ensemble -> 
  Télécharger (JavaScript Notation)

  grade_to_genote.py -t 'xls_group' -s ./INF808-AL_groups.json -g 1 -G 2 -f 2 -l 39 ./Presentations_lundi.xlsx  ./notes-IFT511Gr1-A2024.xlsx
  grade_to_genote.py --type 'xls_group' --moodle_json_group ./INF808-AL_groups.json --group_col 1 --grade_col 2 --row_header 2 --last_row 39 ./Presentations_lundi.xlsx  ./notes-IFT511Gr1-A2024.xlsx


  """
  parser = argparse.ArgumentParser( prog='grade_to_genote', description=description, usage=usage )
  parser.add_argument( 'grades',  type=pathlib.Path, nargs=1,\
    help="Imported Grades file (mandatory)" )  
  parser.add_argument( 'genote',  type=pathlib.Path, nargs=1,\
    help="Genote file (mandatory)" )
  parser.add_argument( '-t', '--type',  nargs='?', default='moodle_xls_report', \
                      choices=[ 'moodle_xls_report', 'score_list', 'moodle_json_exam', 
                          'xls_group' ],\
    help="""Type of file being imported. 
    1) score_list: indicates grades will be imported from a JSON 
      Score List File.
    2) moodle_xls_report: indicates grades will be imported from a
      moodle XLS Report File. While the file may contain the grades
      from multiple exams, we assume the file only contains a single
      exam. This means that during the exportation from 
      moodle, a single exam has been selected.
    3) moodle_json_exam: indicates the grades will be imported from 
      a Moodle JSON Exam file
    4) 'xls_group' : indicates that group and corresponding grades 
      are imported from an 'home made' XLS file. As a result Group
      and Grade columns need to be clearly pointed with the 
      --group_col, --grade_col, --row_header, and --last_row 
      parameters. In addition, composition of the groups are provided
      by the Moodle JSON Group File indicated with the 
      --moodle_json_group parameter.

    """ )
  parser.add_argument( '-g', '--group_col',  type=int, nargs='?', \
    help="""Indicates the column index of the Group column in the XLS file. 
    Index count starts from 0. The XLS File contains a Group and a 
    corresponding Grade column.""" )
  parser.add_argument( '-G', '--grade_col',  type=int, nargs='?', \
    help="""Indicates the column index of the Grade column in the XLS file. 
    Index count starts from 0. The XLS File contains a Group and a 
    corresponding Grade column.""" )
  parser.add_argument( '-f', '--row_header',  type=int, nargs='?', \
    help="""Indicates the row index of the header of the Group and Grade 
    columns in the XLS file. Index count starts from 0. The XLS File 
    contains a Group and a corresponding Grade column.""" )
  parser.add_argument( '-l', '--last_row',  type=int, nargs='?', \
    help="""Indicates the row index of the last row of the Group and Grade
    columns in the XLS file. Index count starts from 0. The XLS File 
    contains a Group and a corresponding Grade column.""" )
  parser.add_argument( '-s', '--moodle_json_group',  type=pathlib.Path, nargs='?', \
    help="""Indicates the path of the Moodle JSON Group File. This file 
    contains the members associated to each group.""" )

  args = parser.parse_args()
#  print( args )

 
  if args.type == 'moodle_xls_report':
    grades = MoodleXLSReportFile(  args.grades[ 0 ] )
    score_list_path = grades.export_score_list( )
  elif args.type == 'moodle_json_exam':   
    grades = MoodleJSONExamFile(  args.grades[ 0 ] )
    score_list_path = grades.export_score_list( )
  elif args.type == 'score_list':
    score_list_path = args.grades[0]     
  elif args.type == 'xls_group':
    for a in [ args.group_col, args.grade_col, args.row_header, args.last_row ]:
      try:
        int( a )
      except ValueError:
        print( f""" --group_col, --grade_col, --row_header,
        --last_row must be provided as index. --moodle_json_group 
        must be provided as a valid path. """ )   
    grades = XLSFile( args.grades[ 0 ] )
    score_list_path = grades.group_export_score_list( moodle_json_group_path=args.moodle_json_group,
                                                     x_col=args.group_col, y_col=args.grade_col,
                                                     row_header=args.row_header, last_row=args.last_row )
  else:
    raise ValueError( f"unknown type {args.type} to import to genote" )  
      
  genote = GenoteXLSFile( args.genote [ 0 ] )
  f = genote.import_score_list( score_list_path )
  print( f"""\nStudent_id and grades have been computed in {f}. To finalize 
  the exportation to the Genote file, please:

    1. open the Genote file {args.genote[0]}
    2. open the newly created file: {f} 
    3. copy the column from the newly created file to the Genote file.
  """ )

