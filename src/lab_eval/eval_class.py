import random
import os
from os.path import join, isdir, isfile, getsize, abspath
import sys
import json
import pyunpack
import shutil
import pathlib
from importlib import import_module, invalidate_caches, reload
import traceback
from cryptography.hazmat.primitives import hashes
import json
import signal, time
import subprocess
import binascii
import math
import time
import datetime 
import argparse
import configparser
import importlib
import pandas as pd

from lab_eval import eval_lab


def create_dir( directory, force=False ):
  """ remove existing dir and creates a clean directory """
  if isdir( directory ) is False:
    os.makedirs( directory )
  else: 
    if len(os.listdir( directory ) ) != 0:
      if force is False:
        raise ValueError( f"{directory} already exists -> re-creating it. "\
                          f"Please delete it. Note you can force the deletion "\
                          f"by setting force to True." )
      #raise ValueError( f" Testing {directory} already exists -> re-creating it."\
      #                f"TBD: activate shutil.rmtree( {directory} )" )
      shutil.rmtree( directory )
      os.makedirs( directory )


class Moodle:

  def extract( self, moodle_zip_file, moodle_dir  ):
    """ created the moodle directory with all students submissions """
    if isfile( moodle_zip_file  ) is False:
      raise ValueError( "Unable to find {moodle_zip_file}" )
    create_dir( moodle_dir )
    pyunpack.Archive( moodle_zip_file ).extractall( moodle_dir )
    
  def set_lab_dir( self, moodle_dir, lab_dir ):
    """ populates lab_dir from moodle_dir

    import considers the various ways a student can submit in moodle
    """
    ## Create the destination directory
    ## Note that import only works with empty destination directory because unziping th emoodle archive creates a directory which contains all student files and the files are then  moved to the expected destination.
    ## moodle.zip --> dst_dir/moodle/<all files>
    ## instead of moodle.zip --> dst_dir/<all files>
    ## We do not see this as an issue and create_dir takes this in charge. 
    ## If that were an issue we could: 
    ## - 1) identify dst_dir/moodle/ in a better way - even unzipping in /tmp. 
    ##  Currently it is identified as the unique directory which implies 
    ##  the initial destination directory is empty.
    ## - 2) see if extract does not hav eth eappropriated feature.
    ## create_dir
    create_dir( lab_dir )
    
    for dir_name in os.listdir( moodle_dir ):
      if isdir( join( moodle_dir, dir_name ) ) is False:
        continue
      #creating the same directory in lab_dir
      dst_dir = join( lab_dir, dir_name )
      os.makedirs( dst_dir )
      src_moodle_dir = join( moodle_dir, dir_name )
      moodle_file_list = os.listdir( src_moodle_dir )
      if len( moodle_file_list ) == 1:
        src_moodle = join( src_moodle_dir, moodle_file_list [ 0 ] )
        if isfile( src_moodle ) and ( '.gzip' in src_moodle or \
                                      '.zip' in src_moodle or \
                                      '.tar' in src_moodle or \
                                      '.gz' in src_moodle or \
                                      '.rar' in src_moodle ): # a archive file gzip, tar, tar.gz, rar...
          pyunpack.Archive( src_moodle ).extractall( dst_dir )
          ## creates an additional directory
          ## moodle.zip --> dst_dir/moodle/<all files>
          ## instead of moodle.zip --> dst_dir/<all files>
          dst_file_list = os.listdir( dst_dir ) 
          unzip_dir = join( dst_dir, dst_file_list[ 0 ] ) 
          if len( dst_file_list ) == 1 and isdir( unzip_dir ):
            for file_name in os.listdir( unzip_dir ) :  
              shutil.move( join( unzip_dir, file_name ), join( dst_dir, file_name ) )
        elif isfile( src_moodle ) and ( '.py' in src_moodle or '.pdf' in src_moodle ):
          shutil.copy( src_moodle ,dst_dir )
        
        elif isdir( src_moodle ): # a directory that contains the files 
          for student_file_name in os.listdir( src_moodle ):
            shutil.copy( join( src_moodle, student_file_name ), 
                       join( dst_dir, student_file_name ) )
        else: 
          raise ValueError( f"{src_moodle} is not a file nor a directory" )
      else: # files are directly provided
        for file_name in moodle_file_list:
          if '.pdf' in file_name or '.py' in file_name:
            shutil.copy( join( src_moodle_dir, file_name ), 
                       join( dst_dir, file_name ) )

  def get_student( self, dir_name, student_db ):
    """ provides a unique id associated to the student given a 
        student directory assigned by moodle 
    """
    dir_name = dir_name.split( ',' )
    last_name = dir_name[ 0 ].replace( '_', ' ' )
    last_name = last_name.strip( )
    reminder = dir_name[ 1 ][ 1 : ].split( '_' )[ : -3 ]
    first_name = reminder[ 0 ]
    if len( reminder ) > 1 :
      for s in reminder [ 1 : ]:
        first_name += f" {s}"
    first_name.strip() 
    return student_db.search( first_name=first_name, last_name=last_name )


class StudentDB:
   
  def __init__( self, json_file_name ):
    """ student DB can be extracted in a json format from moodle """
    self.json_file_name = json_file_name

    if isfile( json_file_name ):
      with open( json_file_name, 'rt', encoding='utf8' ) as f:
       ## the list of student  extracted by moodle is actually a list of list
       ## [[ {student_1}, {student_2},... ]] so we take the the first element.
        self.db = json.loads( f.read( ) )[ 0 ]
       ## check the lables are expected and otherwise raise an error so the 
       ## user change sthe labels. 
    else:
      raise ValueError( f"""Cannot instantiate StudentDB json_file_name
      {json_file_name} is not a file. """ )

  def search( self, first_name=None, last_name=None ):
    for student in self.db :
      ## uds has french label. I think the best way is to standardize the
      ## JSON file to default value and detect if the scheme is respected.
      ##if student[ 'firstname' ] == first_name and student[ 'lastname' ] == last_name:
      if student[ 'prnom' ] == first_name and student[ 'nomdefamille' ] == last_name:
        return student
    raise ValueError( f"Student with firstname {first_name} and lastname"\
                      f" {last_name} not found." )



class EvalClass:

  def __init__( self, lab_conf, class_dir, student_db=None, moodle=None ): 
    """ evaluates the labs 

    Args:
      - lab_conf: the file path configuration file
        to evaluate a single lab.
      - class_dir: the directory associated to the class,
        that is the directory where all students labs and 
        evaluations are expected to be located.
      - student_db: the database of the students. The data
        base is mostly intended to correlate the submitted 
        files to some student associated information. Currenlty 
        we use that data base to name the logs with the 
        student id. 

    This class assumes the following structure:
    +- class_dir
      +- grade_list.json 
      +- moodle_dir
          +- student1-firstname__last_name
          +- student2-firstname__last_name
          +- student3-firstname__last_name
          +- student3-firstname__last_name
      +- lab_dir
        +- student1-firstname_last_name_id
        +- student2-firstname_last_name_id
        +- student3-firstname_last_name_id
        +- student3-firstname_last_name_id
      +- log_dir
        +- student1-firstname_last_name_id.log
    """
    
    
    self.lab_conf = os.path.abspath( os.path.expanduser( lab_conf ) )
    self.class_dir = os.path.abspath( os.path.expanduser( class_dir ) )
    config = configparser.ConfigParser()
    config.read( self.lab_conf )
    try:
      self.json_score_list = os.path.abspath( os.path.expanduser(\
        config[ 'Instructor' ][ 'json_score_list' ] ) )
    except KeyError:
      self.json_score_list = os.path.join( self.class_dir, 'score_list.json' )

    self.lab_dir = join( self.class_dir, 'lab' )
    try:
      self.log_dir = os.path.abspath( os.path.expanduser(\
        config[ 'Instructor' ][ 'log_dir' ] ) ) 
    except KeyError:
      self.log_dir = join( self.class_dir, 'log_dir' )
    for directory in [ self.lab_dir, self.log_dir ]:
      if isdir( directory ) is False:
        os.makedirs( directory )
    self.student_db = student_db
    self.moodle = moodle

  def get_lab_id( self, dir_name ):
    """ returns the lab_id

    The lab_id is derived from the directory associated to the student. 
    In our case we use the student id as the lab_id when possible.
    This requires that both self.moodle and self.student_db are provided. 

    Args:
      - dir_name: the directory associated to the student
    
    Returns:
      - lab_id (str): the student username.
    """
    if self.student_db is None or self.moodle is None:
      lab_id = dir_name
    else:
      student = self.moodle.get_student( dir_name, self.student_db )
##      lab_id = student[ 'username' ]
      lab_id = student[ 'nomdutilisateur' ]
    return lab_id




  def eval_class( self, force=False ):
    if isfile( self.json_score_list ) is True:
      if force is True :
        raise ValueError( f"{self.json_score_list}: os.remove( self.json_score_list )" ) 
      else:
        response = input( f"Grades have already been recorded in the file "\
                          f"{self.json_score_list}.\n"\
                          f"To remove and re-run the evaluation type R.\n"\
                          f"Otherwise the evaluation will skip evaluation "\
                          f"aleary performed." )
        if response == 'R':
          raise ValueError( f"{self.json_score_list}: os.remove( self.json_score_list )" ) 

    if isdir( self.lab_dir ) is False:
      raise ValueError( f"Cannot perform lab evaluation lab_dir "\
        f"{self.lab_dir} not found." )
    for dir_name in os.listdir( self.lab_dir ):
      student_lab_dir = str( join( self.lab_dir, dir_name ) )
      student_id = str( self.get_lab_id( dir_name ) )
      student_log_dir = str( join( self.log_dir, dir_name ) )
      if isdir( student_log_dir ) is False:
        os.makedirs( student_log_dir )
      print( f"Evaluating {student_lab_dir} for {student_id}, log dir {student_log_dir}" )
      subprocess.run( [ "lab_eval_lab",  "--conf",  str( self.lab_conf ),\
                        "--lab_id", student_id, "--log_dir", student_log_dir,\
                        "--json_score_list", str( self.json_score_list ), student_lab_dir ] )
    log_archive = str( join( join( self.log_dir, os.pardir ), os.path.basename( self.log_dir ) ) )
    shutil.make_archive( log_archive, 'zip', str( self.log_dir ) )

  def detect_same_files( self,  file_name_list=[] ):
    """ reports identical files with same hash 

    Args:
      file_name_list: contains a list of files that should be looked at. 
      By default all files are considered whic may include input files 
      present in multiple if not all directories. Such files are not subject 
      to duplicate detection.  
    """
    hash_dict = {} 
    for student_dir in os.listdir( self.lab_dir ):
      student_file_list = os.listdir( join( self.lab_dir, student_dir) )
      if file_name_list == []:
        file_name_list = student_file_list
      for file_name in student_file_list:
        if file_name not in file_name_list:
          continue
        with open( join( self.lab_dir, student_dir, file_name), 'rb' ) as f:
          digest = hashes.Hash(hashes.SHA256())
          digest.update( f.read() )
          h = digest.finalize()
          meta_data = { 'name' : student_dir, 'file' : file_name } 
          if h in hash_dict.keys():
            hash_dict[ h ].append( meta_data )
          else: 
            hash_dict[ h ] = [ meta_data ]
    ## removing hash with no collision
    for h in list( hash_dict.keys() ):
      if len( hash_dict[ h ] ) == 1 :
        del hash_dict[ h ]
    if len( hash_dict.keys() ) == 0:
      print( "No collisions detected: all files seems different" )
    else:
      print( "Some collisions between files have been detected" )
      print_hash_dict = {}
      for k in hash_dict.keys() :
        print_hash_dict[ f"{binascii.hexlify( k, sep=' ' )}"  ] = hash_dict[ k ] 
      print( json.dumps( print_hash_dict, indent=2 ) )
        



class ScoreList:

  def __init__( self, json_score_list, eval_lab_obj=None ):
    """ provides means to manipulate the json_score_list 

    Args:
      - json_score_list (path): designate sthe files that represents
      the scores.
      - eval_lab_obj: the instantiation of the specific Eval<lab_name>Lab
      class. We need this class to update the grades as this class contains
      the marking scheme.   
    """
    self.json_score_list = json_score_list
    self.eval_lab_obj = eval_lab_obj
    if eval_lab_obj is None :
      self.eval_lab_obj = eval_lab.EvalLab( json_score_list=json_score_list )

  def init_from_moodle_json( self, moodle_json_path, \
#                  json_score_list_path='./json_score_list.json',
                      max_total_score=None ):
  
    """ initializes score_list from a moodle score list format

    When a test is performed on moddle, we can export the results in a json file, 
    but this resulting file does not have the format of the score_list. 
    This function initiates a ScoreList object.

    Args:
      moodle_json_path (str): the path of the moodle json file
      json_score_list_path (str): the file of the output json_score_list_path.
        By default, the output file is './json_score_list.json'.
      max_total_score (int): the maximum score. Thisis usefull to derive 
        the grade as as percentage. By default the highest score is considered.  
    """

    json_score_list = {}

    with open( moodle_json_path, 'rt', encoding='utf8' ) as f:
      moodle_list = json.loads( f.read( ) )[ 0 ]
     
    for moodle_student in moodle_list:
      question_dict = {}
      question_index = 1
      for k in moodle_student.keys() :
        if k[0] == 'q':
          ## moodle by default writtes numbers like "2,3" instead of "2.3"
          ### we ensure the format is appropriated
          ## In some cases the value is also "-" which is replaced by 0
# print( f"moodle_student: {moodle_student}" )
          v = moodle_student[ k ]
          if v == '-':
            v = 0
          else:
            v = v.replace( ',', '.' )
          question_dict[ question_index ] = float( v )
          question_index += 1
      ## the current format seems to have changed    
      ## if len( question_dict.keys() ) == 0:    
      json_score_list[ moodle_student[ 'nomdutilisateur' ] ] = question_dict

    self.record_score_list( json_score_list  )
    self.finalize( max_total_score=max_total_score )

  def init_from_file( json_score_list_path, conf_path=None ):
    """ initialize Scorelist from conf and json file_name

    Args:
      json_score_list_path (path): file name of the json score list file.
      conf_path (path): file path of the configuration file. This is necessary to 
        generate the EvaluationClass and get appropriated information to compute
        the grades. If not specified, it is set to None. 
    """

    json_score_list = os.path.abspath( os.path.expanduser( json_score_list_path ) ) 
  
    config = configparser.ConfigParser()
#    print( f"::: conf_path: {conf_path}" )
#    print( f"::: conf_path:expanduser {os.path.expanduser( conf_path )}" )
#    print( f"::: conf_path:abs(expanduser) {os.path.abspath( os.path.expanduser( conf_path ))}" )
    config.read( os.path.abspath( os.path.expanduser( conf_path ) ) )
  
    ## getting element that describe the class object to evaluate
    ## the various instances of the labs.
    ## we need this class to compute the grades 
    eval_class = config['LabEvaluationClass']['eval_class']
    module = config['LabEvaluationClass']['module']
    module_dir = os.path.abspath( os.path.expanduser( \
                   config['LabEvaluationClass']['module_dir'] ))
    ### import the class the represents the evaluation (from conf )
    ## The following lines intend to instantiate the module
    ## lab_caesar.EvalCaesarLab lab_babyesp.EvalBabyESP
    ## The module_dir is added to the path so the python
    ## interpreter can find the module
    ## The designated class is instantiated from the module
    sys.path.insert(0, os.path.abspath( os.path.expanduser( module_dir ) ) )
#    print( f"::: sys.path: {sys.path}" )
    specific_lab_module = importlib.import_module( module )
    specific_lab_class = getattr( specific_lab_module, eval_class )
  
#    self.__init__( json_score_list, specific_lab_class() )
    return ScoreList( json_score_list, specific_lab_class() )

  def load_score_list( self, input_file=None ) :
    """ load the JSON file and returns the content 

    """
    if input_file is None:
      input_file = self.json_score_list
    with open( input_file, 'r', encoding='utf8' ) as f:
      content = json.loads( f.read() )
    return content

  def record_score_list( self, grade_list, output_file=None ):
    if output_file is None:
      output_file = self.json_score_list
    with open( output_file, 'w', encoding='utf8' ) as f:
      f.write( json.dumps( grade_list, indent=2 ) )


  def add_score_list(self, added_file, initial_file=None, output_file=None ):
    """adds added_file to initial_file and exported the result in output_file.

    This is expected to be used, when the reports are evaluated 
    separately from the scripts. This leads to two different 
    score_list files that need to be merged. 
    """
    if initial_file is None:
      initial_file = self.json_score_list
    add_content = self.load_score_list( input_file=added_file )
#    print( f"::: add_content: {add_content.keys()}" )
    initial_content = self.load_score_list( input_file=initial_file )
    for lab_id in list( add_content.keys() ):
#      print( f"::: BEFORE initial_content : {initial_content[ lab_id ]}" )
      for question in add_content[ lab_id ].keys():
        if isinstance( add_content[ lab_id ][ question ], ( float, int) ):
          if initial_content[ lab_id ][ question ] is None:
            initial_content[ lab_id ][ question ] = 0
          initial_content[ lab_id ][ question ] += add_content[ lab_id ][ question ]
        else:
          print( f" >>> {lab_id}[ {question} ] : {add_content[ lab_id ][ question ]} not added" )
#      print( f"::: AFTER initial_content : {initial_content[ lab_id ]}" )
#print( f"::: initial_content: {initial_content}" )
    self.record_score_list( initial_content, output_file=output_file )
    
    

  def get_min_max_mean_grade( self, max_total_score=None ):
    """ determines the min, max and mean total grades

    Here "total grade" means the sum of the scores obtained
    for each questions. In other words the total grade is an
    addition of teh score and is not normalized. 
    """
    score_list = self.load_score_list( )

    max_grade = 0
    min_grade = 0
    mean_grade = 0
    for lab_id in list( score_list.keys() ):
      ## print( f"lab_id: {lab_id}" )
      ## score = score_list[ lab_id ] 
      #lab = self.eval_lab_class( join( self.lab_dir, dir_name), \
      ## In order to provide a percentage or a grade, we need
      ## the scoring scheme. This is the reason we initialize a 
      ## lab
#      print( f"type( self.eval_lab_obj ) : {type( self.eval_lab_obj )}" )
      self.eval_lab_obj.score = score_list[ lab_id ]         
#      lab = self.eval_lab_class( None, None, \
#                          self.instructor_dir, \
#                          json_score_list=self.json_score_list, \
#                          lab_id=lab_id )
#      lab.compute_grade( )
      self.eval_lab_obj.compute_grade( max_total_score=max_total_score )
#      print( f" self.eval_lab_obj.compute_grade( ) : {self.eval_lab_obj.score }" )
      score_list[ lab_id ] = self.eval_lab_obj.score  
      total_grade = self.eval_lab_obj.score[ "grade (total)" ]
      max_grade = max( total_grade, max_grade ) 
      min_grade = min( total_grade, min_grade ) 
      mean_grade += total_grade
    self.record_score_list( score_list )
    mean_grade /= len( score_list )
    return min_grade, max_grade, mean_grade



  def finalize( self, h_mean=3.2/5*100, h_std=1, max_total_score=None ):
    """Finalizes the scores and provides an estimation so grades have the
       specified mean value.

    Args:
      max_total_score (int): indicates the maximum possible score. This is usefull to compute the grade in %. The value is considered by eval_lab.EvalLab.compute_score(). We need to define this value as the eval_lab class does not have the grade_scheme being set.  In the future, we should probably hav ethis variable in the instantiation of the class or as as function.   
    """
    min_grade, max_grade, mean_grade = self.get_min_max_mean_grade( max_total_score=max_total_score )
    print( f"min: {min_grade}" )
    print( f"max: {max_grade}" )
    print( f"mean: {mean_grade}" )
    grade_list = self.load_score_list( )
    std = 0
    std_perc = 0
    mean_grade_perc = 0
    for lab_id in list( grade_list.keys() ):
      grade = grade_list[ lab_id ] 
      ## When grade (%) has already been computed, we do not re-compute it.
      ## The reason is that for labs, for example, it should be computed 
      ## by the lab_eval class.
      ## When grade (%) has not been computed, when max_total_score is 
      ## provided this value is used to determine teh percenatge otherwise 
      ## we use max value of all scores.
      ## print( f"grade.keys() : {grade.keys()}" )
      if "grade (%)" not in list( grade.keys() ) :
        if max_total_score is not None:
          grade[ "grade (%)" ] = grade[ "grade (total)" ] / max_total_score * 100.0
        else:
          grade[ "grade (%)" ] = grade[ "grade (total)" ] / max_grade * 100.0

      mean_grade_perc += grade[ "grade (%)" ]
      std += grade[ "grade (total)" ] ** 2
      std_perc += grade[ "grade (%)" ] ** 2
      grade_list[ lab_id ] = grade

    n = len ( list( grade_list.keys() ) )
    mean_grade_perc /= n
    std -= mean_grade ** 2
    std_perc -= mean_grade_perc ** 2
    std = math.sqrt( std )
    std_perc = math.sqrt( std_perc )
    if n != 1: 
      std /=  n - 1
      std_perc /= n - 1

    for lab_id in list( grade_list.keys() ):
      if std != 0:
        grade = grade_list[ lab_id ]
        grade[ "h grade (total)" ] = h_std / std * ( grade[ "grade (total)" ] - mean_grade ) + h_mean
        grade[ "h grade (%)" ] = h_std / std * ( grade[ "grade (%)" ] - mean_grade_perc ) + h_mean
        grade_list[ lab_id ] = grade
    self.record_score_list( grade_list )

#  def set_xls_grades( self, lab_id, score_list ):
#    try:
#      return score_list[ lab_id ][ "grade (%)" ]
#    except KeyError:
#      return lab_id

  def export_xls( self, xls_file, student_id_row=17, student_id_col=0, sheet_name="grade" ):
    """ complete the xls sheet with the grades

    xls sheet contains a columns of student_ids which starts at student_id_row,
   student_id_column. 
   This function extracts the column, computes the grades and then fill the grades
   in grade_columns in the orifignal file. 

   Note that we have not been able to find a way to complete the existing sheet, 
   so instead we just create a new sheet.

    Args:
      xls_file (path): the file from which students ids are read and grades are completed.
      student_id_row (int): the row where the first students id is 
      student_id_col (int): the column where the first student_id is
      grade_col (int): the column where the grades are placed.
    """

    score_list = self.load_score_list( )
#    print( score_list.keys() )
#    print( 'casb1901' in score_list.keys() )
#    print( score_list[ 'casb1901' ] )
    def get_score( student_id ):
      if student_id in score_list.keys(): 
        grade = score_list[ student_id ][ "grade (%)" ] 
      else:
        grade = 0
      return grade
##    ##df = pd.read_excel('notes-INF808Gr18-A2023.xlsx', engine='openpyxl', usecols='A')
##    ##df = pd.read_excel( xls_file, engine='openpyxl', headers=15 )
##    ## get the list of student student_row, student_col
    ## the engine needs to be specified to read xls (at leats the new format)
    ## we need to read -2 as the first row being read is interpreted as the header
    ## in our case the header contains the name o fthe column and is overwritten 
    df = pd.read_excel( xls_file, engine='openpyxl', header=student_id_row - 2,\
                        usecols=[ student_id_col ], names=[ "students_id" ] )
    ## removing NaN
    df = df.dropna( )
#    print( f"df: {df}" )
#    print( "----" )
#    print( f"{df.dtypes}" )
#    print( "----" )
#    print( f"{df.head}" )
#    print( "----" )
    #df[ "grade (%)" ] = df[ "student_id" ].apply( lambda x: score_list[ str(x) ][ "grade (%)" ], axis=1 ) 
    
    df[ "grade (%)" ] = df[ "students_id" ].apply(  get_score ) 
    
##    ## df[ destination_col ] = df[ student_col_id].apply( lambda x: self.set_xls_grades( x, score_list ) ) 
## df.to_excel( xls_file, engine='openpyxl' )
    df.to_excel( './grade_' + sheet_name + '.xls', engine='openpyxl' )
#    with pd.ExcelWriter(xls_file, mode="a", engine="openpyxl") as writer:
#      df.to_excel( writer, engine='openpyxl', sheet_name=sheet_name )
#        ## header=None, \
#        ## columns=[ "student_id", "grade (%)" ], index=False sheet_name=sheet_name)
#                ## sheetstartrow=student_id_row - 1, \
#                ## startcol=grade_col, index=False )   
#      writer.save()

  def notify_students( self, student_db, project_name=""):
    """lists the students for which an evaluation has been performed

    The evaluation is performed remotely, and results are sent to the
    students. This function collects the emails and prepares the mail
    to be sent. In the future, it would be good the email is sent 
    automatically.   
    """
    score_list = self.load_score_list( )
    student_list = []
    for student in student_db.db:
      ## The student list considers students in db that have 
      ## their lab evaluated.  
      ##if student[ 'username' ] in score_list.keys():
      if student[ 'nomdutilisateur' ] in score_list.keys():
        student_list.append( student )
    l_student = len( student_list )
    l_lab = len( score_list.keys() )
#    if len( student_list ) != len( score_list.keys() ):
    if l_student != l_lab :
      print( "ERROR: student_list and score_list.keys() mismatches\n" )
      if l_student > l_lab:
        print( """It is likely that the student database contains 
          multiple times the same student ID""" )
      if l_lab > l_student:
        print( """It is likely that students are missing in the DB""" )
      print("""\n
          - score_list contains the evaluation of the students. It is a 
        dictionnary where the key is the student ID and the the value 
        contains the various grades of the questions. \n
          - student_list is a list of students that have been evaluated. 
        It is mostly a list of students that have been evaluated. \n 
        In our case, we expect both list to hav ethe same length as one 
        lab corresponds to one student. This is not what we have found.\n\n
          - len( student_list ): {len( student_list )} 
          - len( score_list.keys() ): {len( score_list.keys() )}""" )

      student_id_list = [ student[ 'nomdutilisateur' ] for student in student_list ]
      intersection = []
      for s in list( score_list.keys() ):
        if s in student_id_list:
          student_id_list.remove( s )
          del score_list[ s ] 
          intersection.append( s )
      print( f"""\n
        - student IDs common to both score_list.keys() and 
          student_list are: {intersection}.
        - student_list : contains the additional students
          associated with the following IDs: {student_id_list}.
        - score_list.keys : contains the additional students
          IDs: {score_list.keys()}""" )
#      for k in score_list.keys():
#        in_db = False  
#        for student in student_db.db:
#          if student[ 'nomdutilisateur' ] == k:
#            inb_db = True
#            break 
#        if in_db is False:
#          print( f"!!! student with id {k} doe snot seem to be in student.db" )
      raise ValueError( f"non matching length student_list" ) 
    
    email_list = ""
    for student in student_list:
##      email_list += f"{student[ 'email' ]}; "
      email_list += f"{student[ 'adressedecourriel' ]}; "
    print( f"\nevaluated students: {email_list}" )
    
    print( f"\n{project_name} scripts {datetime.datetime.now()}" ) 
    print( "\nBonjour,\n" )
    print( "Veuillez trouver ci-joint les résultats des scripts soumis." )
    print( "Nous rappelons que les soumissions sont individuelles.\n" )
    print( "Cordialement" )
    print( "Daniel" )

def lab_eval_class():

  description = \
  """ create and overwritte the project 

  This is usually the step performed upon receiving scripts from the students.
  This steps is usually followed by an manual correction of evaluation of the
  grades - typically from a report. 
  """


  parser = argparse.ArgumentParser( description=description )
  parser.add_argument( 'moodle_zip_file',  type=pathlib.Path, nargs=1,\
    help="correpsonds to the archive that contains all labs from\
         students. (mandatory)" )
  parser.add_argument( '-conf', '--conf',  required=True, \
    type=pathlib.Path, nargs=1,\
    help="configuration file (mandatory)")
  parser.add_argument( '-student_db', '--student_db', \
    type=pathlib.Path, nargs='?', default=None,
    help="file of the student data base (optional). This \
          argument overwrite the one specified in the conf file,\
          if specified in the configuration file. When student_db \
          is neither specified in the configuration file or in \
          the command line. It is set by default to None.")
  parser.add_argument( '-class_dir', '--class_dir',  default='./class_dir', \
    type=pathlib.Path, nargs=1,\
    help="directory with all files and evaluation file (mandatory)")
  args = parser.parse_args()

  ## initializing moodle_zip_file
  moodle_zip_file =  os.path.abspath( os.path.expanduser( \
              args.moodle_zip_file[ 0 ] ) )
  
  ## intializing the confile
  conf =  os.path.abspath( os.path.expanduser(  args.conf[ 0 ] ) )
  config = configparser.ConfigParser()
  config.read( conf )

  ## initializinf student_db
  if args.student_db is not None:
    json_student_db = os.path.abspath( os.path.expanduser( args.student_db ) )
    student_db_src = "CLI"
  else:
    try:
      print( f"student_db: from config file: {config[ 'ClassEvaluation' ][ 'student_db' ]}" )
      json_student_db = os.path.abspath( os.path.expanduser( \
                  config[ 'ClassEvaluation' ][ 'student_db' ] ) )
      print( f"student_db: after expansion {json_student_db}" )
      student_db_src = "conf file"
    except KeyError:
      json_student_db = None
      student_db_src = "not provided"
  if json_student_db is not None:
    print( json_student_db  )
    student_db = StudentDB( json_student_db )
    print( student_db.db )
  else:
    student_db = None  
  ## intialization of class_dir
  class_dir = os.path.abspath( os.path.expanduser( args.class_dir[ 0 ] ) ) 

   
  if isfile( moodle_zip_file ) is False:
    raise ValueError( f"Unexpected {moodle_zip_file}\n Expecting file." )
  moodle_dir = os.path.join( class_dir, 'moodle_dir' )
  create_dir( class_dir, force=True )

  print( f""" 
  - moodle_zip_file: containing the labs submitted by
    the students: {moodle_zip_file}
  - conf: configuration file for lab evaluation and 
    class evaluation ({student_db_src}): {json_student_db}
  - class_dir: directory that will contain labs, and
    associated logs: {class_dir}.""")

  moodle = Moodle( )
  moodle.extract( moodle_zip_file, moodle_dir )
  eval_class = EvalClass ( conf, class_dir, student_db=student_db, moodle=moodle )
  print( f"""
  - log_dir: conatins the resulting evaluations of the labs.
    {eval_class.log_dir}. If that value is not expected.
    You can do the following:
    * 1) consider the default value 'log_dir', and comment 
      the corresponding value in the conf file. 
    * 2) configure the conf file with the expected value.
  - json_score_list: designates the file that contains the scores:
    {eval_class.json_score_list}. If that is not the expected 
    value please see log_dir comment to find out how to proceed.""")

  moodle.set_lab_dir( moodle_dir, eval_class.lab_dir )
  eval_class.detect_same_files()
  eval_class.eval_class( )
  ## preparing the email to send
  score_list = ScoreList( eval_class.json_score_list )
  project_name = os.path.dirname( class_dir )
  score_list.notify_students( student_db, project_name=project_name)

## init_from file should be moved to ScoreList
## 
###def init_score_list( conf_path, json_score_list_path ):
###
###  json_score_list = os.path.abspath( os.path.expanduser( json_score_list_path ) ) 
###
###  config = configparser.ConfigParser()
###  config.read( os.path.abspath( os.path.expanduser( conf_path ) ) )
###
###  ## getting element that describe the class object to evaluate
###  ## the various instances of the labs.
###  ## we need this class to compute the grades 
###  eval_class = config['LabEvaluationClass']['eval_class']
###  module = config['LabEvaluationClass']['module']
###  module_dir = os.path.abspath( os.path.expanduser( \
###                 config['LabEvaluationClass']['module_dir'] ))
###  ### import the class the represents the evaluation (from conf )
###  ## The following lines intend to instantiate the module
###  ## lab_caesar.EvalCaesarLab lab_babyesp.EvalBabyESP
###  ## The module_dir is added to the path so the python
###  ## interpreter can find the module
###  ## The designated class is instantiated from the module
###  sys.path.insert(0, os.path.abspath( module_dir ) )
###  specific_lab_module = importlib.import_module( module )
###  specific_lab_class = getattr( specific_lab_module, eval_class )
###
###  return ScoreList( json_score_list, specific_lab_class() )

def lab_finalize_grades( ):
  description = """ computes grades 

  This includes total as well as compute harmonized grades.
  This is especially usefull when some questions have been 
  evaluated manually or when the scores need to be harmonized.

  """
  parser = argparse.ArgumentParser( description=description )
  parser.add_argument( '-conf', '--conf',  required=True, \
    type=pathlib.Path, nargs=1,\
    help="configuration file (mandatory)")
  parser.add_argument( 'json_score_list',  type=pathlib.Path, nargs=1,\
    help="specifies the score list file (mandatory)" )
  args = parser.parse_args()

  score_list = ScoreList.init_from_file( args.json_score_list[ 0 ], args.conf[ 0 ] )

##  json_score_list = os.path.abspath( os.path.expanduser( args.json_score_list[ 0 ] ) ) 
##
##  config = configparser.ConfigParser()
##  config.read( os.path.abspath( os.path.expanduser(  args.conf[ 0 ] ) ) )
##
##  ## getting element that describe the class object to evaluate
##  ## the various instances of the labs.
##  ## we need this class to compute the grades 
##  eval_class = config['LabEvaluationClass']['eval_class']
##  module = config['LabEvaluationClass']['module']
##  module_dir = os.path.abspath( os.path.expanduser( \
##                 config['LabEvaluationClass']['module_dir'] ))
##  ### import the class the represents the evaluation (from conf )
##  ## The following lines intend to instantiate the module
##  ## lab_caesar.EvalCaesarLab lab_babyesp.EvalBabyESP
##  ## The module_dir is added to the path so the python
##  ## interpreter can find the module
##  ## The designated class is instantiated from the module
##  sys.path.insert(0, os.path.abspath( module_dir ) )
##  specific_lab_module = importb1	lib.import_module( module )
##  specific_lab_class = getattr( specific_lab_module, eval_class )
##
##  score_list = ScoreList( json_score_list, specific_lab_class() )


  score_list.finalize( )

def lab_add_score_list():

  description = """ adds an external json_score_list to the cureent one

  This is typically usefull when different json_score_list for exemple 
  one generated by the evaluation of the scripts and one generated for 
  the report need to be combined.  
  """
#  def complete_xls( self, xls_file, student_id_row=17, student_id_col=0, grade_col=4 ):
  parser = argparse.ArgumentParser( description=description )
  parser.add_argument( '-conf', '--conf',  required=True, \
    type=pathlib.Path, nargs=1,\
    help="configuration file (mandatory)")
  parser.add_argument( 'json_score_list',  type=pathlib.Path, nargs=1,\
    help="specifies the score list file (mandatory)" )
  parser.add_argument( 'added_json_score_list',  type=pathlib.Path, nargs=1, \
    help="file to be added.  (mandatory)" )
  args = parser.parse_args()

  score_list = ScoreList.init_from_file( args.json_score_list[ 0 ], args.conf[ 0 ] )
  score_list.add_score_list( args.added_json_score_list[ 0 ] )
  score_list.finalize( )

def lab_export_xls():

  description = """ exports the labs grades to a pre-filled xls file

  This is usually useful to import the grades to the system.
  """
  parser = argparse.ArgumentParser( description=description )
  parser.add_argument( '-conf', '--conf', \
    type=pathlib.Path, nargs='?', default=None, const=None,\
    help="configuration file (mandatory)")
  parser.add_argument( 'json_score_list',  type=pathlib.Path, nargs=1,\
    help="specifies the score list file (mandatory)" )
  parser.add_argument( 'xls_file',  type=pathlib.Path, nargs=1,\
    help="file to be completed. This is the file from where students ids are read.  (mandatory)" )
  parser.add_argument( '-student_id_row', '--student_id_row',  type=int, nargs='?',\
    default=17, const=0, help="row index where student id starts" )
  parser.add_argument( '-student_id_col', '--student_id_col',  type=int, nargs='?',\
    default=0, const=0, help="row index where student id starts" )
  parser.add_argument( '-sheet_name', '--sheet_name',  type=ascii, nargs='?', default="grades",\
    const="grades", help="name of the shee that contains the grades" )
#  parser.add_argument( '-max_total_score', '--max_total_score',  type=int, nargs='?',\
#    default=None, const=None, help="indicates the max_total_score (optional). By default it takes the  max 'observed score'" )
#  parser.add_argument( '-grade_col', '--grade_col',  type=int, nargs='?', default=4,\
#    const=4, help="column index where grades are written." )

  args = parser.parse_args()
  print( f"args: {args}" )
  if args.conf is None :
    score_list =  ScoreList( args.json_score_list[ 0 ] )
  else:
    ## We need here to initialize the Scorelist and then run init_from_file
    ## the function nneds to be added the 'self' argument.
    score_list = ScoreList.init_from_file( args.json_score_list[ 0 ], args.conf )
  ## we initially finalize to ensure the json_score_list remains coherent. 
  ## we believe this shoudl be handled outside the exportation.
  ## if so one needs to provide the necessary arguments to finalize.
  ## score_list.finalize( )
  ## self.finalize( max_total_score=max_total_score )

  score_list.export_xls( args.xls_file[ 0 ], student_id_row=args.student_id_row,\
                           student_id_col=args.student_id_col,\
                           sheet_name=args.sheet_name[1:-1] )

#def moodle_json_to_score_list( moodle_json_path, \
#                              json_score_list_path='./json_score_list.json',
#                              max_total_score=None ):

def moodle_json_to_score_list( ):
  """ put the scores returned by moodle to a score_lits format

      we want to take advanatge of what has been developped with 
      score_list. Especially with finalize and export_xls facilities. 
      Exportation to xls is possible directly from moodle, but ther eare 
      some decimal convesion that makes it hard to handle. 
  """

  description = """convert moodle_json scores to the score_list format"""
  parser = argparse.ArgumentParser( description=description )
  parser.add_argument( 'moodle_json_score_list',  type=pathlib.Path, nargs=1,\
    help="specifies the moodle score list file (mandatory)" )
  parser.add_argument( '-json_score_list', '--json_score_list', \
    type=pathlib.Path, nargs='?', default='./score_list.json',\
    const='./score_list.json', \
    help="specifies the score list file (optional). By default ./json_score_list" )
  parser.add_argument( '-max_total_score', '--max_total_score',  type=int, nargs='?',\
    default=None, const=None, help="indicates the max_total_score (optional). By default it takes the  max 'observed score'" )

  args = parser.parse_args()
  print( args )
  score_list = ScoreList( args.json_score_list )
  score_list.init_from_moodle_json( args.moodle_json_score_list[ 0 ], \
     max_total_score=args.max_total_score )
   




  
#  score_list = ScoreList( json_score_list, specific_lab_class() )
#  score_list.finalize( )
#  score_list.xls( xls_file, student_id_col='A' ):

## if __name__ == "__main__":
  
#  lab_eval_class() 
#  lab_finalize_grades()    
