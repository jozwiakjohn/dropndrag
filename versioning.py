#!/usr/bin/env python

# 2015 Oct 7, John Jozwiak for Apple RF Systems.

import glob,os,re,traceback

###################################################################################

class MultiMap(object):

  def  add( self , x , y ):
    if x not in self.d:
      (self.d)[x] = set()
    (self.d)[x].add(y)

  def __init__( self , list_of_pairs ):
    self.d = {}
    try:
      for (x,y) in list_of_pairs:
        self.add( x , y )
    except:
      print '# DEBUG MultiMap init usage:',list_of_pairs

  def  remove( self , x , y ):
    if x in self.d and y in (self.d)[x]:
      (self.d)[x].remove(y)

  def  map( self , x , y ):
    if x in self.d:
      return (self.d)[x]
    else:
      return None

  def  structure( self ):
    return self.d

###################################################################################

def  git( path , command , verbose = False ):

  cmnd = 'git -C ' + path + ' ' + command
  if verbose:
    print '## ',cmnd
  rslt = os.popen( cmnd ).read()

  rslt = rslt.strip().split('\n')
  if verbose:
    for line in rslt:  #  rslt[:3] :
      print '#> ', line

  return rslt

###################################################################################

def  initialize( path ):

  print '\n#### DEBUG versioning.initialize(',path,')\n'

  git( path , 'init'                                           , verbose = False)
  for f in glob.glob( os.path.join( path , '*' ) ):
    git( path , 'add ' + f , verbose = False)
  git( path , 'config user.email "ptablecrew@group.apple.com"' , verbose = False)
  git( path , 'config user.name  "ptables"'                    , verbose = False)
  git( path , 'commit -m "initial check-in."'                  , verbose = False)

###################################################################################

def  add_commit( doc , as_user='noone' ):

  from Tool_for_Excel import filepath_for_egg
  eggroot = filepath_for_egg( doc )
  rslt = git( eggroot , 'status' )

  mods = []
  for line in rslt :
    if 'modified' in line :
      line = line.split(':')[1]
      line = line.strip()
      mods.append( line )

  for line in mods :
    git( eggroot , 'add ' + line )

  message = as_user + ' : '
  for line in mods:
    message += line + ' , '

  git( eggroot , 'commit -m "' + message + '"' )

###################################################################################

def  do_history ( doc ): # does not track a specific doc at the moment.

  if 'None' == doc :
    return []

  from Tool_for_Excel import docs_root_get

  doc_root = os.path.join( docs_root_get() , doc )

  history = git( doc_root , 'log --oneline' )
  revs    = []

  for line in history :
    line = line.strip()
    line = re.sub(' [ ]*',' ',line)
    line = line.split(' ',1)
    if len(line) == 2:
      revs.append( line )

  return revs

###################################################################################

def  versions_available_for_doc( doc ):
  h = do_history( doc )
  return h

###################################################################################

def  most_recent_commit_for_doc( doc ):
  h = do_history( doc )
  result = h[0][0]
  return result

###################################################################################

def  most_recent_commit_for_doc_by_user( doc , username ):

  history = do_history( doc )
  #  For example,
  #  history = [
  #              ['fd52539', 'jozwiak : Indian/E1 , Indian/F1 ,']
  #            , ['6f2dbbd', 'jozwiak : Indian/B1 ,']
  #            , ['3235039', 'jozwiak : Indian/F4 ,']
  #            , ['0efa50b', 'jozwiak : Indian/D2 ,']
  #            , ['c1aae3c', 'initial check-in.']  
  #            ]

  for event in history:
    if username in event[1]:
      return event[0]
  
  return None

###################################################################################

def  do_flashback( doc , hexstring ):  #  commandline's "flashback hexstring".
  print '\n# FLASHBACK start',doc,hexstring
  from Tool_for_Excel import docs_root_get
  doc_root = os.path.join( docs_root_get() , doc )

  git( doc_root , 'branch --list' )
  git( doc_root , 'branch futz ' + hexstring )
  git( doc_root , 'branch --list' )
  git( doc_root , 'checkout futz' )

  print '# flashback done, to ', hexstring , '\n'

###################################################################################

def  do_back_to_now( doc ):  #  commandline's "now".
  print '\n# BACK TO NOW start',doc
  from Tool_for_Excel import docs_root_get
  doc_root = os.path.join( docs_root_get() , doc )

  git( doc_root , 'checkout     master' )
  git( doc_root , 'branch   -D  futz'   )
  git( doc_root , 'branch   --list'     )
  
  print '# back to now, done.\n'

###################################################################################

def  paths_changed_over_timespan( doc , earlier , later , verbose = False ):  #  returns a list of paths

  from Tool_for_Excel import docs_root_get
  doc_root = os.path.join( docs_root_get() , doc )
  history  = do_history( doc )

  if 'latest' == later:
    later = history[0][0]

  if verbose:
    print '# full history for doc',doc,':'
    for line in history:
      print '    ',line
 
  # consider verifying here that later happens after earlier, meaning later happens before earlier in the list, ironically.
  just_rev_nums = [ h[0] for h in history ]
  print '# just rev nums =',just_rev_nums

  if ( (earlier not in just_rev_nums) or
       (later   not in just_rev_nums) or
       (just_rev_nums.index(later) >= just_rev_nums.index(earlier))
     ):
    print '# warning : paths-changed-over-timespan was given earlier and later revs either out of order or not both valid.'
    print '#           earliar =',earlier,' and later =',later
    return MultiMap([]).structure()

  # at this point, both rev nums exist and are in proper order.

  since = []
  begun = False or ('latest' == later)
  for h in history:
    if later       in h[0]:
      begun = True 
    if not begun:
      continue
    if earlier not in h[0]:
      since.append(h)
    else:
      break

  if verbose:
    print '# shorter history for doc',doc,':'
    for line in since:
      print '    ',line
 
  since = [ h[1] for h in since ]
  since = [ re.sub( '^.*: ' , '' , h ) for h in since ]
  since = [ re.sub( ' ,'    , '' , h ) for h in since ]
  since = [ h.strip()                  for h in since ]
  since = [ h.split(' ')               for h in since ]

  edits = []
  for commit in since:  #  h is a list of paths, each of the form 'Indian/D4', corresponding to a commit.
    edits = edits + commit

  #  each edit looks like 'Indian/E1' now, so break each such to a pair like ['Indian','E1'].

  edits = [ e.strip().split('/') for e in edits ]
  edits = MultiMap( edits )

  if verbose:
    print '# tersest history for doc',doc,':',edits.structure()
 
  return edits.structure()

###################################################################################

def  do_diffs_latest_vs_older( doc , older ):  #  returns a list of paths

  return paths_changed_over_timespan( doc , older , 'latest' )

###################################################################################
