#!/usr/bin/python

###################################################################################
# 2014 Nov 18 - 2015 Oct, John Jozwiak for Apple RF WETA (Tiberiu and Mohit and Prasanna)
###################################################################################

import BaseHTTPServer , ssl
import urlparse

import cgi
import copy
import json  # for its pretty printer.
import re
import os
import socket  #  to get the fully qualified domain name, fqdn.
import time
import uuid
import glob
import Tool_for_Excel
import approvals
import session
import ldap
import projects
import sys
sys.path.append('bottle-0.12.8')
import bottle  #  a single file web framework.
import powertable_workflow
import file_dependence
import versioning

###################################################################################
appleconnecttest     = 'https://wdg2.apple.com/ssowebapp/login.jsp?appID=2744&languageCode=US-EN'
howtohttps           = 'http://www.piware.de/2011/01/creating-an-https-server-in-python/'
Apple_Open_Directory = 'https://istweb.apple.com/mac/server/server_od.html'
###################################################################################
# Notes:
#  convert to python 3.5 in summer 2015 on OSX?
#
#  consider javascript rewrite...preparatory to eventual native app with restful services.
#
#  2015.09.20 : (work for at home days) consider bottlepy.org for code cleanup, and angularjs.org .
#
#  board      limits :  Loui Sanguinetti's group.
#  regulatory limits :  Manjit and friends.
#  sar        limits :  Jon King and DJ.
###################################################################################

def  display_html_header():
  return  \
'''<html>

  <!--
         copyright 2014 January 5 jozwiak@apple.com 
         for Apple RF with TB Muresan, Prasanna Thiagarajan, Mohit Narang.
  -->

  <head>
    <meta http-equiv="cache-control" content="no-cache" />

    <style type="text/css">

      input           {
                        border:none               ;
                        padding: 0 0 0 0          ;
                        border-collapse: collapse ;
                        font-size:7pt             ;
                      }

      table, tr, td   {
                        border-collapse: collapse ;
                        border:none               ;
                      }

      body *          {
                        font-family : sans-serif  ;
                        font-size   : 10px        ;
                      }

      body.white      {
                        background-color : white  ;
                        color            : black  ;
                      }

      body.black      {
                        background-color : black  ;
                        color            : green  ;
                      }

      span.spaceyspan {
                        margin-right: 20px        ;
                      }

      li              {
                        list-style: none          ;
                      }
      a , a:visited , a:hover , a:active { color: inherit;
                                         }
    </style>

<!--<script src="eggshell.js"></script>-->
    <script src="angularjs.min.js"></script>

    <script>
            var  cells_to_save = '';

            function change_trap ( what )
            {
               cells_to_save = cells_to_save + '&' + what.name + '=' + what.value;
               //alert(cells_to_save);
               what.style.backgroundColor = "yellow";
            }

            function save_edited_cells ( server )
            {
               earl = server + cells_to_save;
               //alert( earl );
               this.document.location.href = earl;
               cells_to_save = '';
            }

            function reload ( server )
            {
               earl = server ; // + cells_to_save;
               //alert( earl );
               this.document.location.href = earl;
               cells_to_save = '';
            }
    </script>
  </head>

  <body>
    <div class="centeredDiv">
       <!--
       <canvas id="canvas" width="800" height="800">
         If you see this, your browser predates 2009, which is absurd considering it is 2015 when jozwiak@apple.com wrote this.
       </canvas>
       -->
    </div>

'''

###################################################################################

def  server_url():
  proto  = 'https://' if (443 == PORT_NUMBER) else 'http://'
  index  = 0
  if 'localhost' in sys.argv:
    index = 1
  host   = [ 'radioflyer.apple.com'
           , 'localhost'
#          , socket.getfqdn()
           ][index]
  server = proto + host + (':%d' % PORT_NUMBER)
  return server

###################################################################################

def  top_of_page( args , debug = False ):
  u = 'nemo'
  if 'username' in args:
    u = args['username']
  output = '<hr/><hr/>[[[ power tables ]]]&nbsp&nbsp&nbsp username : %s &nbsp&nbsp&nbsp&nbsp&nbsp<a href="%s">Log Out</a><hr/>' % (u,server_url()+'/logout?username='+u)
  if debug:
    for a in args:
      output += str(a) + ':' + str(args[a]) + '<br/>\n'
    output += '<hr/>'
  return output

###################################################################################

def  colophon():

  output = ''
  output += '<hr/><hr/><center><a href="mailto:jozwiak@apple.com?subject=\'Feature Request for Web Excel\'">Request a Feature</a><br/><br/><span style="font-size:10px">&copy; Copyright 2015 John Jozwiak (jozwiak@apple.com) for Apple.</span></center>'
  output += '<hr/><hr/>'

  return output

###################################################################################

def  login_page_html( args ):

  output = '<hr/><hr/>'
  output += ('\n\n<form action="%s/%s" method=post>' % (server_url() , 'select_document' ) )
  output += '  [[[ power tables ]]]\n'
  output += '  <input type=text      name="username"  value="username">\n'
  output += '  <input type=password  name="password"  value="password">\n'
  output += '  <input type=submit>\n'
  output += '  <input type=reset>\n'
  output += '  <span style="font-size:13px">&nbsp;&nbsp;( Open Directory credentials )&nbsp;&nbsp;</span>'
  output += '  </form>\n'
  output += '<hr/><hr/>'

  return output

###################################################################################

def  href( text , server , slash_to_question_mark , args_map ):
  o = '<a href="' + str(server) + '/' + slash_to_question_mark
  if len(args_map) > 0:
    o += '?'
    for x in args_map:  #  as in x -> y in the map. 
      o += (str(x) + '=' + str(args_map[x]) + '&')
  o += ('">' + text + '</a>')
  return o

###################################################################################

active_when = 'latest'  #  these should hold hashes, since comments may duplicate.

###################################################################################

def  show_when( doc , tab , username , when , vs ):

  verbose = False
  if verbose:
    print '# show when : doc      =',doc
    print '# show when : tab      =',tab
    print '# show when : username =',username
    print '# show when : when     =',when
    print '# show when : vs       =',vs
 
  vs = [ ['latest','latest'] ] + vs

  global active_when

  output = ''

  for choice in vs:  #  [ ['f84f4b0', 'Belgium/B1 was changed'], ['c4acafa', 'initial check-in.'] ]

    if choice[0] == active_when:
      output += ' [ '

    output += ( '<a href="' + server_url() + '/edit_document?' + 'doc=' + doc           + '&'
                                                               + 'tab=' + tab           + '&'
                                                               + 'username=' + username + '&'
                                                               + 'when=' + choice[0]    + '">'
                + choice[1] + '</a>'
              )

    if choice[0] == active_when:
      output += ' ] '

    output += '<option value="' + choice[0] + '"' + (' selected' if choice[0] == when else '') + '>' + choice[1] + '</option>'
    output += '&nbsp&nbsp&nbsp'

  output += '</select>'
  output += '</form>'

  if verbose:
    print '# show when : return output=',output

  return output

###################################################################################

def  html_form_begin( func , args , method = 'get' ):  #  or method = 'post'

  output = '<form style="display:inline" method="' + method + '" action="' + (server_url() + '/' + func) + '"'

  arrr  = Tool_for_Excel.debug_printable_copy_of_dict(args)

  if 'get' == method:  #  build the ?a=b&c=d sort of suffix.

    first = True
    for a in arrr :
      if first:
        first = False
        output += '?'
      else:
        output += '&'
      output += (a + '=' + arrr[a])
    output += '>'

  elif 'post' == method:
    output += '>'
    for a in arrr:
      output += '<input type=hidden name="' + a + '" value="' + arrr[a] + '"/>\n'
  
  print '# html_form_begin = ',output
  return output

###################################################################################

def  dropdown_of_available_eggshells( name = 'dropdown' , pattern = None , prepend = None ):

  output =  '<select name="' + name + '">'
  docs = Tool_for_Excel.available_loaded_docs()

  if (prepend):
    docs = [prepend] + docs

  for d in docs:
    output += ( '<option value="' + d + '">' + d + '</option>' )  #  need to fix bug in dropdowns here.
  output += '</select>'
  return output

###################################################################################

def  dropdown_of_available_versions_for_doc( doc = None ):

  if (doc) :
    verz = versioning.versions_available_for_doc( doc )
  else:
    verz = []

  print '# DROPDOWN debug : verz =',verz,'for doc =',doc

  output =  '<select name="selected_version">'
  for v in range(len(verz)) :
    output += ( '<option value="' + str(verz[v][0]) + '">' + str(verz[v][1]) + '</option>' )
  output += '</select>'
  return output

###################################################################################

def  dropdown_of_available_projects( prepend = '' ):
  # this sorts the list of projects available, and prepends a '*' to mean all.
  output =  '<select name="selected_project">'
  projs  = projects.available()
  projs  = sorted(projs)
  if prepend != ''   :
    projs  = [prepend] + projs

  fltr = projects.filter_get()

  for d in projs:
    yep = ( (fltr.lower() == d.lower()) and ( prepend != '___') )
    output += ( '<option value="' + d + '"' + (' selected ' if yep else '') + '>' + d + '</option>' )
  output += '</select>'
  return output

###################################################################################

def  dropdown_of_ptablecrew_approvers():
  output    = '<select name="selected_approver">'
  approvers = powertable_workflow.open_directory_ptablecrew_approvers()
  for a in approvers:
    output += ( '<option value="' + a + '">' + a + '</option>' )
  output += '</select>'
  return output

###################################################################################

def  dropdown_of_ptablecrew_dris():
  output    = '<select name="selected_dri">'
  dris      = powertable_workflow.open_directory_ptablecrew_dris()
  for d in dris:
    output += ( '<option value="' + d + '">' + d + '</option>' )
  output += '</select>'
  return output

###################################################################################

def  main_menu( args ):

  username = 'nemo_displaylistofloadedexceldocuments'
  when     = 'latest'
  wanorlan = 'wwan'
  if 'username' in args:  username = args['username']
  if 'when'     in args:  when     = args['when'    ]
  if 'wanorlan' in args and 'wlan' == args['wanorlan']:
    pass #  if 'wlan' then do combine and generate a new file and add and commit the new file.

  loaded_eggshells_exist = ( [] != Tool_for_Excel.available_loaded_docs() )

  output = display_html_header()
  output += top_of_page( args )

  workbooks = Tool_for_Excel.do_list()
  workbooks = [ w[0] for w in workbooks ]
  workbooks = sorted(workbooks)

  ### Project list UI

  if True :  #  [] != workbooks :

    output +=  '<hr/><br/>Show only Projects with &nbsp;&nbsp;'
    output += ('<form onchange="submit()" style="display:inline" enctype="multipart/form-data" action="%s/%s" method="POST">' % (server_url() , 'project_filter') )
    output += dropdown_of_available_projects( prepend = '(anything)' ) + '&nbsp; in their names.<br/><br/>'
    output +=    '<input type=hidden name="username" value="' + username + '"/>'
    output +=  '</form>'

  if True :  #  [] != workbooks :

    output += 'Add a new Project with name '
    output += ('<form style="display:inline" enctype="multipart/form-data" action="%s/%s" method="POST">' % (server_url() , 'project_name_add') )
    output +=    '<input type="text"   name="project_name" style="outline:solid 1px" />'
    output +=    '<input type=hidden name="username" value="' + username + '"/>'
    output +=  '</form>'

    if len(projects.available()) > 0:
      output +=  ', or<br/><br/>Delete a Project (but not associated files) named '
      output += ('<form onchange="submit()" style="display:inline" enctype="multipart/form-data" action="%s/%s" method="POST">' % (server_url() , 'project_name_delete') )
      output += dropdown_of_available_projects( prepend = '___' )
      output +=    '<input type=hidden name="username" value="' + username + '"/>'
      output +=  '</form>'
      output +=  '<br/><br/>'

  ### Work with an existing server document.

  def  _proj ( project__role ):
    return project__role.split('__',1)[0]

  ###

  def  _role ( project__role ):
    if '__' not in project__role:
      return project__role
    return project__role.split('__',1)[1]

  ###

  def  ename ( project__role ):  #  excel file's name that was loaded to create this eggshell.
    e = powertable_workflow.excel_name( project__role )
    return e

  ###

  def  _dri  ( project__role ):  #  excel file's name that was loaded to create this eggshell.
    e = powertable_workflow.excel_name( project__role )
    return e

  ###

  if [] != workbooks :
    output += '<hr/><br/>Work with an already uploaded Excel workbook...<br/><br/>'
    output += '<table style="border-collapse:collapse; border-spacing:0; border:none;cellpadding=0;cellspacing=0;">'

    output += '<tr>'
    output += '<td>Project</td>'
    output += '<td>Name / Project Role</td>'
    output += '<td>Original Name</td>'
    output += '<td>Excel download</td>'
    output += '<td>Erase</td>'
    output += '</tr>'

    for w in workbooks :
      print '# WORKBOOK',w
      output += '<tr>'
      output += '<td>' + _proj(w) + '</td>'
      output += '<td><a href="%s">%s</a></td>' % (server_url()+'/edit_document?doc='+str(w) + '&tab=None&username=' + username + '&when=' + when , _role(w))
      output += '<td>%s</td>' % (ename(w))
    # output += '<div  style="">'
      output += '<td><span style="font-family:Times;font-size:18px;">( <a href="' + (server_url() + '/download?doc=' + w + '&username=' + username)  + '">download as an Excel .xls file</a> )</span></td>'
      output += '<td><span style="font-family:Times;font-size:18px;">( <a href="' + (server_url() + '/erase?doc='    + w + '&username=' + username)  + '">erase from server</a> )</span></td>'
    # output += '<td><span style="font-family:Times;font-size:18px;">( ' + html_form_begin( 'rename' , args ) + '<input style="submit" value="mitsub"/><input style="outline:solid 1px" type=text value="new name"/></form> )</span></td>'
    # output += '<td><input type=text></input></td>'
      output += '</tr>'

    output += '</table>\n'
    output += '<br/>\n'
    output += '<br/>'

  #   See links below.
  #        http://www.w3.org/TR/html-markup/input.file.html .
  #        http://stackoverflow.com/questions/8659808/how-does-http-file-upload-work .

  ### Cellular project UI  :  rse , sar , conducted , tolerances

  output +=  '<hr/><br/>Define/Upload a new Cellular Project...<br/><br/>'
  output += ('<form enctype="multipart/form-data" action="%s/%s" method="POST">' % (server_url() , 'sanity_check_uploads') )
  output +=  '<table style="border-collapse:collapse; border-spacing:0; border:none; cellpadding=0; cellspacing=0;">'
  output +=    '<tr><td>project</td><td>:</td><td>' + dropdown_of_available_projects() + '</td></tr>'
  output +=    '<tr><td>dri    </td><td>:</td><td>' + dropdown_of_ptablecrew_dris() + '</td></tr>'

  output +=    '<tr><td> conducted  </td><td>:</td>'
  if loaded_eggshells_exist : output +=      '<td>%s</td><td>or</td>'  %  dropdown_of_available_eggshells( prepend = '-' , name ='dropdown_conducted' )
  output +=        '<td><input type=file name="uploadedfilename_conducted"  /></td></tr>'

  output +=    '<tr><td> rse        </td><td>:</td>'
  if loaded_eggshells_exist : output +=      '<td>%s</td><td>or</td>'  %  dropdown_of_available_eggshells( prepend = '-' , name ='dropdown_rse' )
  output +=        '<td><input type=file name="uploadedfilename_rse"        /></td></tr>'

  output +=    '<tr><td> sar        </td><td>:</td>'
  if loaded_eggshells_exist : output +=      '<td>%s</td><td>or</td>'  %  dropdown_of_available_eggshells( prepend = '-' , name ='dropdown_sar' )
  output +=        '<td><input type=file name="uploadedfilename_sarcellular"/></td></tr>'

  output +=    '<tr><td> tolerances </td><td>:</td>'
  if loaded_eggshells_exist : output +=      '<td>%s</td><td>or</td>'  %  dropdown_of_available_eggshells( prepend = '-' , name ='dropdown_tolerances' )
  output +=        '<td><input type=file name="uploadedfilename_tolerances" /></td></tr>'

  output +=  '</table><p/>'
  output +=  'Name the Auto-updated derived Master File&nbsp:&nbsp'
  output +=    '<input type="text"   name="requested_eggname" style="outline:solid 1px" /><br/><br/>'
  output +=    '<input type=submit value="Use These Files"/>'
  output +=    '<input type=hidden name="username" value="' + username + '"/>'
  output +=    '<input type=hidden name="wanorlan" value="cellular"/>'
  output +=  '</form>'

  ### WiFi project UI  :  BandEdge , BoardLimits , Normal , SAR , Skeleton

  output +=  '<hr/><br/>Define/Upload a new WiFi Project...<br/><br/>'
  output += ('<form enctype="multipart/form-data" action="%s/%s" method="POST">' % (server_url() , 'sanity_check_uploads') )
  output +=  '<table style="border-collapse:collapse; border-spacing:0; border:none; cellpadding=0; cellspacing=0;">'
  output +=    '<tr><td>project</td><td>:</td><td>' + dropdown_of_available_projects() + '</td></tr>'
  output +=    '<tr><td>dri    </td><td>:</td><td>' + dropdown_of_ptablecrew_dris() + '</td></tr>'

  output +=    '<tr><td> skeleton      </td><td>:</td>'
  if loaded_eggshells_exist : output +=      '<td>%s</td><td>or</td>'  %  dropdown_of_available_eggshells( prepend = '-' , name ='dropdown_skeleton' )
  output +=        '<td><input type="file"   name="uploadedfilename_skeleton"/></td></tr>'

  output +=    '<tr><td> Band Edge     </td><td>:</td>'
  if loaded_eggshells_exist : output +=      '<td>%s</td><td>or</td>'  %  dropdown_of_available_eggshells( prepend = '-' , name ='dropdown_bandedge' )
  output +=        '<td><input type="file"   name="uploadedfilename_bandedge"/></td></tr>'

  output +=    '<tr><td> Board Limits  </td><td>:</td>'
  if loaded_eggshells_exist : output +=      '<td>%s</td><td>or</td>'  %  dropdown_of_available_eggshells( prepend = '-' , name ='dropdown_boardlimits' )
  output +=        '<td><input type="file"   name="uploadedfilename_boardlimits"/></td></tr>'

  output +=    '<tr><td> PPSD          </td><td>:</td>'
  if loaded_eggshells_exist : output +=      '<td>%s</td><td>or</td>'  %  dropdown_of_available_eggshells( prepend = '-' , name ='dropdown_ppsd' )
  output +=        '<td><input type="file"   name="uploadedfilename_ppsd"/></td></tr>'

  output +=    '<tr><td> SAR           </td><td>:</td>'
  if loaded_eggshells_exist : output +=      '<td>%s</td><td>or</td>'  %  dropdown_of_available_eggshells( prepend = '-' , name ='dropdown_sar' )
  output +=        '<td><input type="file"   name="uploadedfilename_sarwifi"/></td></tr>'

  output +=  '</table><p/>'
  output +=  'Name the Auto-updated derived Master File&nbsp:&nbsp'
  output +=    '<input type="text"   name="requested_eggname" style="outline:solid 1px" /><br/><br/>'
  output +=    '<input type="submit" value="Use These Files"/>'
  output +=    '<input type="hidden" name="username" value="' + username + '"/>'
  output +=    '<input type="hidden" name="wanorlan" value="wifi"/>'
  output +=  '</form>'

  ### Standalone File Upload UI

  output +=  '<hr/><br/>Upload a new individual Excel file (and possibly rename it)...<br/><br/>'
  output += ('<form enctype="multipart/form-data" action="%s/%s" method="POST">' % (server_url() , 'individual_excel_upload') )
  output +=  '<input type=file name="uploadedfilename_singleton"  /><p/>'
  output +=  'Name the individually uploaded file&nbsp:&nbsp'
  output +=    '<input type="text"   name="requested_eggname" style="outline:solid 1px" /><br/><br/>'
  output +=    '<input type=submit value="Upload it"/>'
  output +=    '<input type=hidden name="username" value="' + username + '"/>'
  output +=  '</form>'

  ### Clone a Document at a Revision to form a New Document

  if loaded_eggshells_exist:

    output +=  '<hr/><br/>Create a New Document by Cloning an Existing Document at a Specified Version...<br/><br/>'
    output += ('<form onchange="submit()" enctype="multipart/form-data" action="%s/%s" method="POST">' % (server_url() , 'clone_promote') )
    output +=  '<table style="border-collapse:collapse; border-spacing:0; border:none; cellpadding=0; cellspacing=0;">'
    output +=    '<tr><td> existing document  </td><td>:</td>'
    output +=        '<td>' + dropdown_of_available_eggshells( prepend = '-' ) + '</td>'
    output +=        '<td>at revision</td>'
    if 'doc' in args:
      tmpdoc = args['doc']
    else:
      tmpdoc = None
    output +=        '<td>' + dropdown_of_available_versions_for_doc( tmpdoc ) + '</td></tr>'
    output +=  '</table>'
    output +=  '<p/>'
    output +=  'Name of the New Document&nbsp:&nbsp'
    output +=  '<input type="text"   name="masterfilename" style="outline:solid 1px" /><br/><br/>'
    output +=  '<input type=submit value="Use This Document at This Revision"/>'
    output +=  '<input type=hidden name="username" value="' + username + '"/>'
    output +=  '</form>'

# ### Sanity check an existing server document.

# output +=  '<hr/><br/>Sanity check an existing document...<p/>'
# output += ('<form enctype="multipart/form-data" action="%s/%s" method="POST">' % (server_url() , 'sanity_check_uploads') )
# output +=     dropdown_of_available_eggshells()
# output +=  '<p/>'
# output +=  '<input type="submit" value="Check that File"/>'
# output +=  '<input type="hidden" name="username" value="' + username + '"/>'
# output +=  '<input type="hidden" name="wanorlan" value="wlan"/>'
# output +=  '</form>'

  ### End of page

  output += colophon()
  output += '</body></html>'
  return output

###################################################################################

def  cell_list_to_html_table( eggpath , sheet , cells , args , changeds ):

  from Tool_for_Excel import cell_read

  ######
  def  editable_cell( coords , value , cellstyle ):  # style is None or a CSS string.
    n     = 'c_' + coords
    style = 'background-color: white'
    if cellstyle :
      style = cellstyle
    return (
             '<input type="text" style="%s" name="%s" value="%s" id="%s" onchange="change_trap(this)"/>'
           ) % (style,n,value,n)
  ######

  rows_set = set()
  cols_set = set()

  for x in cells:
    r = re.sub( '[A-Z]*' , '' , x )
    c = re.sub( '[0-9]*' , '' , x )

    rows_set.add(r)
    cols_set.add(c)

  rows = list(rows_set)
  rows = [ int(x) for x in rows ]
  rows = sorted( rows )
  cols = sorted(list(cols_set))

  is_derived = file_dependence.is_a_derived_doc( args['doc'] )

  output = ''
  action = server_url() + '/edit_document'
  funky  = action + '?doc=' + args['doc'] + '&tab=' + args['tab'] + '&username=' + args['username']
# output += '<hr/>' + action + '<hr/>'

  if not is_derived:
    output += '<input  type="submit"  value="Save to Server"                 onclick="save_edited_cells(\'' + funky + '\');"/>'
  else:
    output += '<input  type="submit"  value="Save Disabled for Derived Docs" onclick="alert(\'This is a derived doc:  edit its precursors.\');"/>'
  output += '&nbsp'
  output += '<input  type="button"  value="Reload from Server"  onclick="reload(\'' + funky + '\');"/>'
  output += '&nbsp'
  output += '<input  type="hidden"  name="doc"      value="' + args['doc']      + '"/>'
  output += '<input  type="hidden"  name="tab"      value="' + args['tab']      + '"/>'
# output += '<input  type="hidden"  name="username" value="' + args['username'] + '"/>'
  output += '<table style="border-style:solid; border-width:1px;">'

  #  the headers of the chart (i.e., the spreadsheet column headers)....

  output += '<tr style="border-style:solid; border-width:1px;">'
  for c in ([''] + cols):  #  e.g., = ['','a','b'].
    output += '<td style="border-style:solid; border-width:1px;">' + c + '</td>'
  output += '</tr>'

  #  the body of the chart....

  for   r in rows:

    output   += '<tr style="border-style:solid; border-width:1px;"><td style="border-style:solid; border-width:1px;">' + str(r) + '</td>'
    for c in cols:

      coords = (c + str(r))
      x      = cell_read ( eggpath , sheet , coords )
      css    = None
      if (coords in changeds):
        css  = 'background-color: orange'
      if is_derived and (x.value != '') and Tool_for_Excel.is_a_Number( x.value ):
        css  = x.style()
        if coords == 'E13':
          print '\n\n\nancestry for E13 =',x.species,'\n\n\n'
      cell_html  = editable_cell( coords ,
                                  x.value , # + str(x.species) ,
                                  css )

      output += '<td style="border-style:solid; border-width:1px;">' + cell_html + '</td>'
    output   += '</tr>'
  output     += '</table>'
  return output

###################################################################################

def  project_tints( p , alldocs ):  #  returns 'w' for wifi or 'c' for cellular.

  pdocs     = filter( lambda x: p in x , alldocs )
  webcolors = powertable_workflow.webcolors
  if p+'__sarcellular' in pdocs:
    t = { 'conducted'   : 'red'    ,
          'rse'         : 'silver' ,
          'sarcellular' : 'yellow' ,
          'tolerances'  : 'lime'
        }
  else:
    t = { 'bandedge'    : 'red'    ,
          'boardlimits' : 'green'  ,
          'sarwifi'     : 'silver' ,
          'ppsd'        : 'orange'
        }
  tt = {}
  for x in t:  tt[x] = webcolors[t[x]]
  return tt

###################################################################################

def  display_eggshell_document( args ):

  root     = server_url()
  doc      = 'None'
  tab      = 'None'
  username = 'noone_displayeggshelldocument'
  when     = 'latest'

  if 'doc'       in args:  doc      = args['doc' ]
  if 'tab'       in args:  tab      = args['tab']
  if 'username'  in args:  username = args['username' ]
  if 'when'      in args:  when     = args['when']
  prj = doc.split('__')[0]

  output = display_html_header()
  output += top_of_page( args )
  output += '<hr/>' + href( 'Main Menu' , root , 'select_document' , {'username':username} ) + '<hr/>'

  ###  list of docs on server

  docs_root = Tool_for_Excel.docs_root_get()
  workbooks = Tool_for_Excel.available_loaded_docs()
  projects  = set( [ pd.split('__')[0] for pd in workbooks ] )

  ###  projects
  output += 'prj : '
  for p in projects :

    if p == prj: output += '<span style="font-size:18px;">['
    output += href( p , root , 'edit_document' , { 'doc':doc , 'tab':'None' , 'username':username , 'when':when , 'prj':p} )
    if p == prj: output += ']</span>'
    output += '&nbsp&nbsp&nbsp'

  ###  docs
  output += '<br/><br/>doc : '
  for w in filter( lambda x : prj in x , workbooks ) :

    if w == doc: output += '<span style="font-size:18px;">['
    output += href( w.split('__')[-1] , root , 'edit_document' , { 'doc':w , 'tab':'None' , 'username':username , 'when':when , 'prj':prj } )
    if w == doc: output += ']</span>'
    output += '&nbsp&nbsp&nbsp'

  docs_root_doc = os.path.join( docs_root , doc )

  ###  tabs of current doc
  output += '<br/><br/>tab : '

  tabs_full_paths = glob.glob( docs_root_doc + '/*' )    #  e.g., [ '/t/e/usa' , '/t/e/sweden' ] .
  tabs = [ x.split('/')[-1] for x in tabs_full_paths ]   #  e.g., [      'usa' ,      'sweden' ] .

  for t in tabs:
    if t == tab: output += '['
    output += href( t , root , 'edit_document' , {'doc':doc , 'tab':t , 'username':username , 'when':when , 'prj':prj} )
    if t == tab: output += ']'
    output += '&nbsp'

  #  versions available for current doc
  versions = versioning.do_history(doc)
  output += '<hr/>version : ' + show_when( doc , tab , username , when , versions ) + '<hr/>'

  #  approvers statuses.
  apprvls = approvals.get( doc )
  o = 'approvers :'
  for person in apprvls:
    if apprvls[person]:
      o += ('<input type=button value="%s" style="background-color:green">' % person) + '</input>'
    else:
      o += ('<input type=button value="%s" style="background-color:red"  >' % person) + '</input>'
    o += '&nbsp;'
  o += '<br/>'
  output += o

  #  coloring (and diffing?) UI

  is_derived = file_dependence.is_a_derived_doc( args['doc'] )

  if not is_derived:
    output += 'cell coloring indicates&hellip;'
    output +=   ' <select name="colors">'
    output +=   '   <option value="rules"    >value versus rules.</option>'
    output +=   '   <option value="newer"    >newer value on server.</option>'
    output +=   '   <option value="original" >original Excel spreadsheet cell color.</option>'
  # output +=   '   <option value="diff"     >diff two tabs at possibly distinct points in time.</option>'
    output +=   ' </select>'

  else:
    output += 'cell coloring indicates originating document for value'
    output += '<table><tr>'
    tints = project_tints( prj , workbooks )
    for t in tints :
      output += ('<td style="' + powertable_workflow.wbtype_to_css(t) + '">' + t + '</td>')
    output += '</tr></table>'

  output +=   '<p/>'

  if (False):  #  display_diff_ui , which depends on "colors" above being "diff".

    # left

    output += '&nbsp;'
    output +=   ' <select name="diff_left">'

    for t in tabs:
      output +=   '   <option value="' + t + '"    >' + t + '</option>'
    output +=   ' </select>'

    output += '&nbsp;@&nbsp;'

    output +=   ' <select name="when_left">'
    for w in versions:
      output +=   '   <option value="' + w[0] + '"    >' + w[1] + '</option>'
    output +=   ' </select>'

    output +=   '&nbsp;versus&nbsp;'

    # right

    output +=   ' <select name="diff_right">'
    for t in tabs:
      output +=   '   <option value="' + t + '"    >' + t + '</option>'
    output +=   ' </select>'

    output += '&nbsp;@&nbsp;'

    output +=   ' <select name="when_right">'
    for w in versions:
      output +=   '   <option value="' + w[0] + '"    >' + w[1] + '</option>'
    output +=   ' </select>'

  output +=   '<hr/>'

  #  debugging the times change delta calculation.

  ltst_global = versioning.most_recent_commit_for_doc        ( doc )
  ltst_users  = versioning.most_recent_commit_for_doc_by_user( doc , username )
  history     = versioning.do_history( doc )
  colorables  = versioning.do_diffs_latest_vs_older( doc , ltst_users )
  changeds    = set()
  if tab in colorables:
    changeds  = colorables[tab]

# output += '<br/>' + doc + '/' + tab + ' for username ' + username + '<br/>'
# output += 'now         : ' + when + '<br/>'
# output += 'ltst_global : ' + str(ltst_global) + '<br/>'
# output += 'ltst_users  : ' + str(ltst_users ) + '<br/>'
# output += 'history     : ' + str(history)     + '<br/>'
# output += 'colorables  : ' + str(colorables)  + '<br/>'
# if tab in colorables:
#   output += 'cells to color : ' + str(colorables[tab])
# output += '<hr/>'

  tab_text = 'No tab is selected yet.'
  if tab != 'None':

      cells    = glob.glob( docs_root_doc + '/' + tab + '/*' )
      cells    = [ x.split('/')[-1] for x in cells ]
      tab_text = cell_list_to_html_table( docs_root_doc , tab , cells , args , changeds )

  output += tab_text

  output += '<br/><br/>' 
  output += colophon()
  output += '</body></html>'
  return output

###################################################################################
# od_ersatz = { 'jozwiak'     : 'x'         ,
#               'thiagarajan' : 'tallmachi' ,
#               'agboh'       : 'indranil'  ,
#             }
# if (u in od_ersatz) and (p == od_ersatz[u]):
###################################################################################

def  password_authenticates_username( u , p ):
  return (
          #('' != p) and
           ldap.authenticates_user( u , p ) and
           ldap.user_is_in_group  ( u , 'radioflyers' )
         )

###################################################################################

def  session_authenticates_username( u , session_in_browser ):

  print '# BEWARE session_authenticates_username is broken.'
  session_in_server = session.get( u )
  simply = ( ( session_in_server != 'None'             ) and
             ( session_in_server == session_in_browser )
           )

  print '# sess auth user (',u,',',session_in_browser,') = ',simply
  if simply:
    return u

  u = session.username( session_in_browser )  # Returns None or a username string.
  print '# sess auth user , second try : u =' , u
  return u

###################################################################################

def  web_transitions( func_and_args ,  #  [ 'function_name' , { 'param0' : value0 , ... }
                      session_in_browser
                    ):  #  returns [ 'html text' , None | new_server_sessionid_string , save_as_file ]

  func = func_and_args[0]
  args = func_and_args[1]

  global active_when

  doc      = 'unknown_doc_in_web_transitions'
  username = 'unknown_username_in_web_transitions'
  password = ''
  when     = 'latest'

  if 'doc'      in args:  doc      = args['doc' ]
  if 'username' in args:  username = args['username']
  if 'password' in args:  password = args['password']
  if 'when'     in args:  when     = args['when']

  print '# web transitions func               =' , func
  print '# web transitions args               =' , Tool_for_Excel.debug_printable_copy_of_dict(args)
  print '# web transitions session in browser =' , session_in_browser
  print '# web transitions session in server  =' , session.get( username )

  new_server_sessionid = None

  ####

  if 'quit' == func :

    global keep_running
    keep_running = False
    return [ 'done' , None , False ]

  ####

  if not session_authenticates_username( username , session_in_browser ) :
    print '# session not authenticated ; try password...'
    if password != '':
      if     password_authenticates_username( username , password ) :
        print '# success with password:  do a resumption?'
        new_server_sessionid  = str(uuid.uuid1())
        session.set( username , new_server_sessionid )
        return  [ main_menu( args ) , new_server_sessionid , False ]
      else:
        print '# password failed\n'
        return  [ login_page_html( {} ) , None , False ] # new_server_sessionid
    else:
      print '# password not given, so irrelevant.'
      return    [ login_page_html( {} ) , None , False ] # new_server_sessionid
  else:
    print '# session authenticated.'

  ####

  if '' == func :

    if session_authenticates_username( username , session_in_browser ):
      s = session.get( username )
      print '# resuming...',json.dumps(s , indent=3),'\n'
      if 'None' == s:
        return     [ login_page_html( args ) , new_server_sessionid , False ]
      else:
        return     [ 'The browser and username from args are authenticated; now resume?' , new_server_sessionid , False ]
    else:
      return       [ login_page_html( args ) , None , False ]

  ####

  elif 'logout' == func:

    session.set( username , 'loggedout' )
    return   [ login_page_html( {}   ) , None , False ]

  ####

  elif 'select_document' == func:

    return   [ main_menu( args ) , new_server_sessionid , False ]

  ####

  elif 'sanity_check_uploads' == func:

    powertable_workflow.derive_derived_eggshell( args )  #  2015 Oct 14 refactoring note from JJ to JJ :  seems should be in Tool for Excel.
    return   [ main_menu( args ) , new_server_sessionid , False ]

  elif 'individual_excel_upload' == func:

    return   [ main_menu( args ) , new_server_sessionid , False ]

  ####

  elif 'project_filter' == func:

    if 'selected_project' in args:
      projects.filter_set( args['selected_project'] )
    return   [ main_menu( args ) , new_server_sessionid , False ]

  ####

  elif 'project_name_add' == func:

    if 'project_name' in args:
      projects.add( args['project_name'] )
    return   [ main_menu( args ) , new_server_sessionid , False ]

  ####

  elif 'project_name_delete' == func:

    if 'selected_project' in args :
      projects.delete( args['selected_project'] )
    return   [ main_menu( args ) , new_server_sessionid , False ]

  ####

  elif 'clone_promote' == func:

    # args = # e.g.,
    # {
    #   'username'          : 'jozwiak'                     ,
    #   'wanorlan'          : 'cellular'                    ,
    #   'selected_eggshell' : 'Tolerences_used_for_program' ,
    #   'version'           : '0'                           ,
    #   'selected_version'  : '0'                           ,
    #   'masterfilename'    : 'doggoneit'
    # }

    if (
         ('selected_eggshell' in args) and
         ('selected_version'  in args) and
         ('masterfilename'    in args)
       ) :
      Tool_for_Excel.do_clone_at_version(args)

    return   [ main_menu( args ) , new_server_sessionid , False ]

  ####

  elif 'edit_document' == func:

    Tool_for_Excel.do_writes( args , username )  #  this records any c_X=value args.

    if when != active_when:
      print '# IN WEB TRANSITIONS active_when =' , active_when , 'and when=', when
      active_when = when
      if 'latest' == active_when:
        versioning.do_back_to_now( doc )
      else:
        versioning.do_flashback  ( doc , active_when )
      args['when'] = active_when

    response = display_eggshell_document ( args )
    return   [ response , new_server_sessionid , False ]

  ####

  elif 'erase' == func:

    if 'doc' in args:
      Tool_for_Excel.do_erase( args['doc'] , username )
    return [ main_menu( args ) , new_server_sessionid , False ]

  ####

# elif 'rename' == func:

#   if 'doc' in args:
#     Tool_for_Excel.do_rename( args['doc'] , username )
#   return [ main_menu( args ) , new_server_sessionid , False ]

  ####

  elif 'download' == func:

    Tool_for_Excel.do_clean_generated_download_files()
    Tool_for_Excel.do_excelify( doc , '/tmp/savethis.xls' )  #  this records any c_X=value args.

    response     = ''
    save_as_file = doc + '.xls'
    if len( glob.glob( '/tmp/savethis.xls' ) ) > 0:
      try:
        with open( '/tmp/savethis.xls' , 'r' ) as f:
          response = f.read()
      except:
        print '# download trouble opening /tmp/savethis.xls'

    return   [ response , new_server_sessionid , save_as_file ]

  ####

  elif (('.js' in func) or ('.html' in func) or ('.png' in func) ):

    nameofafile = func
    try:
      with open( nameofafile , 'r' ) as f:
        response = f.read()
      return   [ response , new_server_sessionid , False ]
    except:
      return   [ ''       , new_server_sessionid , False ]

  ####

  elif func in [ 'set_filename'   , 'get_filename' ,
                 'set_session'    , 'get_session'  ,
                 'get_cell'       , 'set_cell'     ,
                 'get_version'    , 'get_revision' ,
                 'commit_version'
               ]:
     return   [ 'performing %s( ... )  ...' % func , new_server_sessionid , False ]
  else:
    return    [ 'unknown command : ' + json.dumps({'func':func,'args':args},indent=3) , new_server_sessionid , False ]

###################################################################################

class EggShellHandler ( BaseHTTPServer.BaseHTTPRequestHandler ) :

   def  parameters( self ):  #  returns [ func , { dict of args to values } ] 

     d    = self.path
     d    = d.strip().split('?')
     d[0] = re.sub('/','',d[0])
     if len(d) < 2:
       d.append( {} )
     else:
       d[1] = d[1].split('&')  #  e.g., 'a=1&b=two' -> [ 'a=1' , 'b=two' ]

     t = {}  #  We compose a dictionary mapping each http argument to its value.

     if 'GET' == self.command:

       for a in d[1]:
         lr = a.split('=')
         # print 'debug c : lr =',lr,' after a =',str(a)
         if 1==len(lr) and type(lr) == type([]):
           lr = lr + ['']
         else:
           pass
           # print '# debug what the heck : type(lr) =',type(lr),' and lr =',lr
         t[lr[0]] = lr[1]

     elif 'POST' == self.command:  #  http://pymotw.com/2/BaseHTTPServer/

       form = cgi.FieldStorage( fp      = self.rfile   ,
                                headers = self.headers , 
                                environ = { 'REQUEST_METHOD' : 'POST' ,
                                            'CONTENT_TYPE'   : self.headers['Content-Type']
                                          }
                              )

       egg_name           = 'EggDoe'
       selected_project   = ''  #  'unspecified_project'

       if 'requested_eggname' in form.keys():
         egg_name         = form['requested_eggname'].value
       if 'selected_project' in form.keys():
         selected_project = form['selected_project'].value

       for field in form.keys():

         t[field] = form[field].value
         print '# FORM FIELD :',field,'=',
         if form[field].filename:

             t[field]         = form[field].filename  #  form[field].value
             excel_file_name  = form[field].filename
             excel_file_bytes = form[field].file.read()
             print '<',len(excel_file_bytes),'bytes of ',excel_file_name,'>'

             Tool_for_Excel.receive_file_from_browser( egg_name         ,
                                                       field            ,
                                                       excel_file_name  ,
                                                       excel_file_bytes ,
                                                       selected_project )  #  i.e., role in project.
             del excel_file_bytes
         else:
             print t[field]

     else:
       print '\n# debug : unexpected http command of',self.command,' (whereas only post and get are used)\n'

     d[1] = t

     print '# http ' , str( self.command) , ' :' , [ d[0] , Tool_for_Excel.debug_printable_copy_of_dict(d[1]) ]
     return d

   #####

   def get_header_as_dict( self ):
     h = str(self.headers).strip().split('\n')  #  a list of lines of form 'X: y\n'
     h = [ x.strip().split(':',1) for x in h ]  #  break each line at the first : so two strings result.
     d = {}
     for [x,y] in h:
       d[x] = y
     return d

   #####

   def get_header_cookies_as_dict( self ):
     cookie_header = self.get_header_as_dict()
     if 'Cookie' in cookie_header:
       cookie_header = self.get_header_as_dict()['Cookie']
       cookies_lines = cookie_header.strip().split(';')  
       cookies_lines = [ x.strip().split('=',1) for x in cookies_lines ] 
     else:
       cookies_lines = []

     cookies_map = {}
     #rint '\n', cookies_lines , '\n'
     for [x,y] in cookies_lines:
      #print len(xy) , xy
       cookies_map[x] = y
     #rint '# DEBUG cookies=' , json.dumps( cookies_map , indent=3 )
     return cookies_map

   #####

   def do_respond( self ):

     print '# self.client_address =',self.client_address
     print '# self.path           =',str(self.path).strip()

     browsers_cookies = self.get_header_cookies_as_dict()
     servers_session  = 'None'
     browsers_session = 'None'
     if 'radioflyersession' in browsers_cookies:
       browsers_session = browsers_cookies['radioflyersession']

     save_as_file = None

     [ r , new_browser_session , save_as_file ] = web_transitions( self.parameters() , browsers_session )

     self.send_response( 200 )

     if not save_as_file :
       self.send_header  ( 'Content-type'        , 'text/html' )
     else:
       self.send_header  ( 'Content-type'        , 'application/octet-stream' )
       self.send_header  ( 'Content-Disposition' , 'attachment; filename="' + save_as_file + '"' )

     if new_browser_session :  #  None if no new session to send to the browser.
       self.send_header  ( 'Set-Cookie'   , 'radioflyersession=' + new_browser_session )

     self.end_headers  ()

     self.wfile.write  ( r )
#    print '\n# type(self.wfile) =',str(type(self.wfile)),'\n'

   #####

   def do_GET (self):  self.do_respond()
   def do_POST(self):  self.do_respond()

   #####

   #    mimetypes = {
   #                  ".html" : 'text/html'              ,
   #                  ".jpg"  : 'image/jpg'              ,
   #                  ".gif"  : 'image/gif'              ,
   #                  ".js"   : 'application/javascript' ,
   #                  ".css"  : 'text/css'               ,
   #                }
   #    mt        = mimetypes['.html']

###################################################################################

if len(sys.argv) > 1:
  if 'port' in sys.argv:
    i_of_port_word = sys.argv.index( 'port' )
    if 1 + i_of_port_word < len(sys.argv):
      PORT_NUMBER = int(sys.argv[ 1 + i_of_port_word ])
  else:
    PORT_NUMBER = 443
else:
  PORT_NUMBER = 443

print '# PORT_NUMBER = %d' % PORT_NUMBER

keep_running = True

try:
   server = BaseHTTPServer.HTTPServer(('', PORT_NUMBER), EggShellHandler)
   if PORT_NUMBER == 443:
     server.socket = ssl.wrap_socket( server.socket , certfile='server.pem' , server_side=True )
   print '# Started BaseHTTPServer.httpserver on port' , PORT_NUMBER

   while keep_running:
     server.handle_request()

except KeyboardInterrupt:
   print '^C received, shutting down the Javascript app server'
   server.socket.close()

###################################################################################
