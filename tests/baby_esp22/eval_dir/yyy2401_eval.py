import sys
sys.path.insert( 0, '/home/emigdan/gitlab/pylab_evaluation/examples' )
import lab_babyesp 

lab = lab_babyesp.EvalBabyESP( '/home/emigdan/gitlab/pylab_evaluation/tests/baby_esp22/lab/lastname2,_firstname2_5696540_assignsubmission_file_', '/home/emigdan/gitlab/pylab_evaluation/tests/test_single_lab/ref', json_score_list='/home/emigdan/gitlab/pylab_evaluation/tests/baby_esp22/lab_grade_list.json', lab_id='yyy2401' )
lab.eval_py() 
