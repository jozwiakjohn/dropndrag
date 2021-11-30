#!/usr/bin/env python

# 2015 Oct 14, John Jozwiak for Apple RF Systems.

import os,traceback,glob,re,sys
from   collections     import  namedtuple
import file_dependence

###################################################################################

# a shape is [
#              eggshell_name                      ,  [0]
#              #tabs                              ,  [1]
#              (c,r) of origin                    ,  [2]
#              dict of sheetname -> (width,height)   [3]
#            ]
# and the shapes_list is sorted such that the skeleton is at the 0 index.
# and c,r are the origin coordinates in shapes_list[0], which is the skeleton.

Shape = namedtuple( 'Shape' , ['docname','ntabs','colsbyrows','sheet_to_dims'] )

###################################################################################

#  https://en.wikipedia.org/wiki/Web_colors
#  I will use WHITE , SILVER , YELLOW , LIME , AQUA , FUCHSIA .

webcolors = {
              'white'   : 'background-color: #FFFFFF' ,
              'silver'  : 'background-color: #C0C0C0' ,
              'yellow'  : 'background-color: #FFFF00' ,
              'lime'    : 'background-color: #00FF00' ,
              'aqua'    : 'background-color: #00FFFF' ,
              'red'     : 'background-color: #CC0000' ,
              'fuchsia' : 'background-color: #FF00FF' ,
              'green'   : 'background-color: #00FF00' , 
              'orange'  : 'background-color: #BB8C60'
            }

wbtype_to_webcolors = {
                        # cellular
                        'conducted'   : 'red'     ,
                        'rse'         : 'silver'  ,
                        'sarcellular' : 'yellow'  ,
                        'tolerances'  : 'lime'    ,

                        # wifi
                        'bandedge'    : 'aqua'    ,
                        'boardlimits' : 'red'    ,
                        'sarwifi'     : 'yellow'  ,
                        'ppsd'        : 'fuchsia'        
                      }

######

def  randomstyle():

  from random import randint

  ri = randint(0,5)
  m  = {
         0 : webcolors[ 'white'   ] ,
         1 : webcolors[ 'silver'  ] ,
         2 : webcolors[ 'yellow'  ] ,
         3 : webcolors[ 'lime'    ] ,
         4 : webcolors[ 'aqua'    ] ,
         5 : webcolors[ 'fuchsia' ]
       }

  return m[ri]

######

def  wbtype_to_css( t ):
  c = webcolors[ 'white' ]
  if (
          (t in wbtype_to_webcolors)
      and (     wbtype_to_webcolors[t] in webcolors)
     ):
    c = webcolors[ wbtype_to_webcolors[t] ] 
  return c
  
######

def  cssstyleforsource( species , demo_cheat = False):

  if demo_cheat:
    return randomstyle()

  s    = webcolors[ 'white'                        ]
  if species in wbtype_to_webcolors:
     s = webcolors[ wbtype_to_webcolors[ species ] ]
  return s

###################################################################################

def  origin_find_virtual( doc ):

  from Tool_for_Excel import map_of_sheet_names_to_sheet_dimensions
  from Tool_for_Excel import cell_read
  from Tool_for_Excel import coords_to_excel_cell_name

  #  Find the Excel cell which holds the string "MIMO device" and whose cell below holds "Channel".
  #  That pair identifies a WiFi fileset upload (a la Shrenik Milapchand's examples).
  #  ..OR..
  #  Find the Excel cell (e.g., "C9") which holds the string "Bands", 
  #  such that the cell to its right (e.g., "D9") holds the string "NV number".
  #  This only tries the upper left 12x12 corner of the doc at its zeroeth sheet;
  #  if there is no success by then, it returns None.
  #  That identifies a Cellular fileset upload (a la Digvijay's examples).

  #####

  def  origin_find_virtual_in_sheet( doc , sheet , dimensions ):

    cell_0_2 = cell_read( doc , sheet , coords_to_excel_cell_name( 0   , 2 ) , whiney=False ).value
    cell_0_3 = cell_read( doc , sheet , coords_to_excel_cell_name( 0   , 3 ) , whiney=False ).value
    if (
         cell_0_2 == 'MIMO device' and
         cell_0_3 == 'Channel'
       ):
      print '# ',doc,'seems to be a WIFI doc, from its detected origin.'
      return (0,0,coords_to_excel_cell_name(0,0))

    (ncols , nrows) = dimensions[ sheet ]
    cbound = min( 12 , ncols )
    rbound = min( 12 , nrows )
    for   c in range(cbound):
      for r in range(rbound):
        if ( 
             ((c+1) <= cbound) and
             ('Bands'     == cell_read( doc , sheet , coords_to_excel_cell_name( c   , r ) , whiney=False ).value ) and
             ('NV number' == cell_read( doc , sheet , coords_to_excel_cell_name( c+1 , r ) , whiney=False ).value )
           ):
            return (c,r,coords_to_excel_cell_name(c,r))
    return None

  #####

  sheets_dims = map_of_sheet_names_to_sheet_dimensions( doc )
  nsheets     = len( sheets_dims.keys() )

  bounds  = (0,0)

  if   ( 0 == nsheets ): 
    return None
  elif ( 1 == nsheets ):
    return origin_find_virtual_in_sheet( doc , sheets_dims.keys()[0] , sheets_dims )
  else:
    if  'Default' in sheets_dims :
      return origin_find_virtual_in_sheet( doc , 'Default' , sheets_dims )
    else:
      return None

###################################################################################

#  class  Origin:  ?  then factor out the import across the two related funcs?

def  origin_locate_and_record( doc ):

  from Tool_for_Excel import docs_root_get

  origin     = origin_find_virtual( doc )
  originfile = os.path.join( docs_root_get() , doc , '.origincoordinates' )

  try:
    with open( originfile , 'w' ) as f:
      f.write( str( origin ) )
  except:
    print '# BUG : locate and record origin\n' ; traceback.print_exc()

###################################################################################

def  origin_lookup( doc ):

  from Tool_for_Excel import docs_root_get

  originfile = os.path.join( docs_root_get() , doc , '.origincoordinates' )

  try:
    with open( originfile , 'r' ) as f:
      try:
        origin = eval(f.read())
      except:
        origin = None

  except:
    print '# BUG : lookup origin\n' ; traceback.print_exc()
    origin = None

  return origin

###################################################################################

def  origin_offsets( list_o_eggshells ):  #  a list of egg names, not paths.
  # Returns: list of [ egg , (offsetc,offsetr) , #sheets , egg_origin_coords , map of sheet to (nc,nr)].
  #          An egg with the most sheets is initial in the returned list.
  from Tool_for_Excel import map_of_sheet_names_to_sheet_dimensions

  #############

  def  delta( originofskeleton , originofother ):  #  where both o1 and o2 are of the form (c,r,excelcellname).
    return  ( originofother[0] - originofskeleton[0] , originofother[1] - originofskeleton[1] )

  #############

  #  A shape is a tuple of the form 
  #        [ eggshell_name , # of sheets , origin , map of sheet_name to (#cols , #rows) ] .

  shapes = [ [p , map_of_sheet_names_to_sheet_dimensions( p )]          for p in list_o_eggshells ]
  shapes = [ [ s[0] , len( s[1].keys() ) , origin_lookup(s[0]) , s[1] ] for s in shapes           ]
  shapes.sort( key = lambda x : x[1] )
  shapes = list(reversed(shapes))

# print '\n# shapes calculation in eggshell offsets...'
# for e in shapes:  #  A shape is [ docname , #sheets , origin_coords , map of sheet to (ncols , nrows)  ].
#   print e
##  print '# ' , e[0] , e[1] , e[2]
##  for s in e[3]:
##    print '#     ',e[3][s],' ',s

  skeleton        = shapes[0]
  skeleton_origin = skeleton[2]
  offsets = [ [ s[0] , delta( skeleton_origin , s[2] ) , s[1], s[2] , s[3] ] for s in shapes ]
  
# for d in offsets:
#   print '# eggshell offset...',d

  return offsets

###################################################################################

def  combine_corresponding_values( sheet , c , r , offsets_map ):  #  needs return species!

  from Tool_for_Excel import is_a_Number , is_a_String , number , cell_read , coords_to_excel_cell_name

  babbly = (sheet == 'Belgium' and coords_to_excel_cell_name(c,r) == 'E13')

  if babbly:
    print '\ncombine corresponding values offsets_map =',offsets_map,'\n'

  #  ...from the point of view of the derived looking back to precursors.
  if type(offsets_map) is list:
    offsets_map = [ [o[0],o[1]] for o in offsets_map ]
    offsets_map = dict(offsets_map)

  if babbly:
    print '\ncombine corresponding values offsets_map =',offsets_map,'\n'

  values  = []
  try:
    for doc in offsets_map:
      (ox,oy) = offsets_map[doc]
      try    :
        v_s = cell_read( doc , sheet , coords_to_excel_cell_name( c+ox , r+oy ) , whiney=False )
        v_s.species = doc.split('__')[-1]
        values.append( v_s )
      except :
        continue
  except :
    print '\nERROR BUG HERE : combine corres',sheet,coords_to_excel_cell_name(c,r),'\n'

  # Now values is a list of Cell instances.
  if babbly and len(values) > 1:
    print 'combine',sheet,coords_to_excel_cell_name(c,r),'values =',values

  ns = filter( lambda c : is_a_Number( c.value )                                , values )
  ss = filter( lambda c : is_a_String( c.value ) and not is_a_Number( c.value ) , values )
  ns.sort( key = lambda x : number(x.value) )
  if babbly and len(ns) > 1:
    print 'DEBUG sorted ns =',ns

  species = None   #  FIX THIS

  if len(ns) > 0 and len(ss) > 0:
    print '# WTF combine corre values : values =' , values , 'ns =' , ns , 'ss =',ss

  if   len(ns) > 0 : return (     ns[0].value , ns[0].species )
  elif len(ss) > 0 : return ( values[0].value , None          )
  else             : return ( values[0].value , None          )

###################################################################################

def  update_derived_cell_from_precursor( doc , sheet , cell , sink_doc , precursors , offsets = None ):

  from file_dependence import dependents_of
  from Tool_for_Excel  import excel_cell_name_to_coords , coords_to_excel_cell_name , cell_write

  verbose = (sheet == 'Belgium' and cell == 'E13') # False
  if verbose:
    print '# FLOW edit of source cell',cell,'in',doc,sheet,'to corresponding cell in',sink_doc

  if sink_doc not in precursors:  #  if the sink_doc has no precursors...there is nothing to do.
    return
  sink_precursors = precursors[ sink_doc ] #  precursors is now a list of strings;  each is a doc name, not path.

  if offsets == None:
    offsets = origin_offsets( sink_precursors )

  offsets = [ [o[0],(-1*o[1][0],-1*o[1][1])] for o in offsets ]  #  now the offset vectors are reversed.
# offsets = filter( lambda x : x[0] in [doc,sink_doc] , offsets )
  offsets = dict( offsets )
  delta   = offsets[ doc ]
  if verbose:
    print '# sink_precursors  :',sink_precursors
    print '# offsets filtered :',offsets
    print '# delta            :',delta

  coords_of_source_cell = excel_cell_name_to_coords( cell )
  coords_of_sink_cell   = ( delta[0] + coords_of_source_cell[0] , delta[1] + coords_of_source_cell[1] )
  (sink_col,sink_row)   = coords_of_sink_cell
  sink_cell             = coords_to_excel_cell_name( coords_of_sink_cell[0] , coords_of_sink_cell[1] )
  if verbose:
    print '# coords src cell  :' , coords_of_source_cell
    print '# coords snk cell  :' , coords_of_sink_cell
    print '# sink_cell        :' , sink_cell

  (value,ancestor) = combine_corresponding_values ( sheet , sink_col , sink_row , offsets )

  if value != None:
    print '\n# UPDATE DERIVED CELL :',sink_doc,sheet,sink_cell,'with value',value,'and ancestor',ancestor,'\n'
    cell_write( sink_doc , sheet , sink_cell , value , species = ancestor , debug=True , commit=False )
  else:
    print '# ignoring update of derived cell with None value : ',sink_doc,sheet,sink_cell

###################################################################################

def  propagate_cell_update_from( doc , sheet , cell ):  #  possibly several dependent master docs.

  from file_dependence import dependents_of

  (precursors,_) = file_dependence.read()  #  precursors is a map from string to list of string.

  for d in file_dependence.dependents_of( doc ):
     update_derived_cell_from_precursor( doc , sheet, cell , d , precursors )

###################################################################################

def  derive_derived_eggshell( args ):

  from Tool_for_Excel import available_loaded_docs , sanitized , write_in_memory_workbook_to_eggshell , Cell
  from Tool_for_Excel import coords_to_excel_cell_name

  #############

  def  combine_eggshells_to_in_memory_wb ( eggs_list ):

    offsets = origin_offsets( eggs_list )
    # offsets is a list, each item of the form
    #   [ egg , (offsetc,offsetr) , #sheets , egg_origin_coords , {s:(nc,nr)...} ],
    # where eggs are ordered by number of sheets, most to fewest.
    skeleton_shape          = offsets[0]
    skeleton_sheets_to_dims = skeleton_shape[4]  #  a map of sheet name to (columns,rows).

    in_mem_wb = {}

    for    sheet  in  skeleton_sheets_to_dims :
      (ncols,nrows) = skeleton_sheets_to_dims[sheet]
      print '# CREATING from skeleton shape',skeleton_shape[0],sheet,'which is',ncols,'x',nrows
      for    col  in  range(ncols):
        for  row  in  range(nrows):

          (value,species) = combine_corresponding_values ( sheet , col , row , offsets )
          assert type(value) is str or type(value) is float or type(value) is int
          if value != None:
            in_mem_wb [ (sheet,coords_to_excel_cell_name(col,row)) ] = Cell(value,species)
            if value != '':
            # print 'in_mem_wb at',sheet,row,col,'=',value,'{',species,'}'
              pass
          else:
            print 'None value at',sheet,row,col

    return in_mem_wb

  #############

  def  clean_role_to_egg_map( role_to_egg_map ):
     eggster = {}
     for e in role_to_egg_map:  #  a map of things like 'ppsd':'Project_blahppsd' .
       t = role_to_egg_map[e]
       t = re.sub( 'uploadedfilename_' , '' , t )
       t = re.sub( 'dropdown_'         , '' , t )
       t = re.sub( '.xlsx'             , '' , t )
       t = re.sub( '.xls'              , '' , t )
       eggster[e] = t
     return eggster

  #############

  print '# derive derived args...'
  soggy = sanitized( args )
  for s in sorted(soggy.keys()):  print '#   ',s,' ---> ',soggy[s]

  if 'selected_project' in args: selected_project = args['selected_project']
  else:                          selected_project = 'NoProject'

  d = args[ 'requested_eggname' ]  #  derived

  # dropdown named files, possibly eclipsed by explicitly named uploaded files.
  if   'wanorlan' not in args:  return
  if   'cellular' == args['wanorlan']:

    c  =  ''                           #  conducted
    r  =  ''                           #  rse
    s  =  ''                           #  sarcellular
    t  =  ''                           #  tolerances

    if         'dropdown_conducted'   in args:  c     = args[         'dropdown_conducted'  ]
    if 'uploadedfilename_conducted'   in args:  c     = args[ 'uploadedfilename_conducted'  ]
    if         'dropdown_rse'         in args:  r     = args[         'dropdown_rse'        ]
    if 'uploadedfilename_rse'         in args:  r     = args[ 'uploadedfilename_rse'        ]
    if         'dropdown_sarcellular' in args:  s     = args[         'dropdown_sarcellular']
    if 'uploadedfilename_sarcellular' in args:  s     = args[ 'uploadedfilename_sarcellular']
    if         'dropdown_tolerances'  in args:  t     = args[         'dropdown_tolerances' ]
    if 'uploadedfilename_tolerances'  in args:  t     = args[ 'uploadedfilename_tolerances' ]

    role2eggmap = { 'conducted':c , 'rse':r , 'sarcellular':s , 'tolerances':t }

  elif 'wifi' == args['wanorlan']:

    sk   =  ''                         #  skeleton
    be   =  ''                         #  bandedge
    bl   =  ''                         #  boardlimits
    ppsd =  ''                         #  ppsd
    sar  =  ''                         #  sar

    if         'dropdown_skeleton'    in args:  sk    = args[         'dropdown_skeleton'    ]
    if 'uploadedfilename_skeleton'    in args:  sk    = args[ 'uploadedfilename_skeleton'    ]
    if         'dropdown_bandedge'    in args:  be    = args[         'dropdown_bandedge'    ]
    if 'uploadedfilename_bandedge'    in args:  be    = args[ 'uploadedfilename_bandedge'    ]
    if         'dropdown_boardlimits' in args:  bl    = args[         'dropdown_boardlimits' ]
    if 'uploadedfilename_boardlimits' in args:  bl    = args[ 'uploadedfilename_boardlimits' ]
    if         'dropdown_ppsd'        in args:  ppsd  = args[         'dropdown_ppsd'        ]
    if 'uploadedfilename_ppsd'        in args:  ppsd  = args[ 'uploadedfilename_ppsd'        ]
    if         'dropdown_sarwifi'     in args:  sar   = args[         'dropdown_sarwifi'     ]
    if 'uploadedfilename_sarwifi'     in args:  sar   = args[ 'uploadedfilename_sarwifi'     ]

    role2eggmap = { 'skeleton':sk , 'bandedge':be , 'boardlimits':bl , 'ppsd':ppsd , 'sarwifi':sar }

  role2eggmap = clean_role_to_egg_map( role2eggmap )
  new_eggs    = [ (selected_project + '__' + x) for x in role2eggmap.keys() ]
  loaded_eggs = available_loaded_docs()  #  assert : new_eggs is a subset of loaded_eggs.
  newname     = selected_project + '__' + d
  file_dependence.add_precursors( newname , new_eggs )

  print '# derive derived : dict_o_eggshells =',role2eggmap,'\nnew_eggs :',new_eggs,'\nloaded_eggs :',loaded_eggs
  print '# derive derived : about to combine ',new_eggs,'to',newname

  in_mem_wb = combine_eggshells_to_in_memory_wb( new_eggs )
  result    = write_in_memory_workbook_to_eggshell( in_mem_wb , newname , original_excel_name = None )

###################################################################################

def  excel_name_record( n ):  #  as of 2015 Nov 15 this is written at line 421 in Tool_for_Excel.
  pass

###################################################################################

def  excel_name( eggshell ):
  from Tool_for_Excel import docs_root_get

  originalnamefilepath = os.path.join( docs_root_get() , eggshell , '.excel_filename' )
  try:
    with open( originalnamefilepath , 'r' ) as h:
      name = h.read()
      name = name.strip()
  except:
      name = '<unknown>'

  return name

###################################################################################

def  open_directory_ptablecrew_approvers():
  import ldap
  return ldap.group_members( 'ptablecrewapprovers' )

###################################################################################

def  open_directory_ptablecrew_dris():
  import ldap
  return ldap.group_members( 'ptablecrewdris' )

###################################################################################
