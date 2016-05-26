from openpyxl import load_workbook
from openpyxl.compat import range
import pygraphviz as pgv 

def change_attributes(graph, nodes_list,shapes_list,attribute_name):
    for x,y in zip(nodes_list,shapes_list):
        if(y==None):
            pass
        else:
            graph.get_node(x).attr[attribute_name]=y

wb = load_workbook(filename='graph.xlsx')
ws = wb.get_sheet_by_name('graph')

parents = []
children = []
parent_shape = []
child_shape = []

for row in ws.rows:
    for cell in row:
        if(cell.row==1):
            pass
        else:
            if(cell.column=='A'):
                    parents.append(cell.value)
            elif(cell.column=='B'):
                    children.append(cell.value)
            elif(cell.column=='C'):
                    parent_shape.append(cell.value)
            elif(cell.column=='D'):
                    child_shape.append(cell.value)
            else:
                pass
            
#make a left-right directed graph (affects 'dot' layout only)
G=pgv.AGraph(rankdir='LR')
#G.graph_attr['fontpath'] = 'C:\\windows\\fonts'


#populate the nodes and edges
for x,y in zip(parents,children):
    #print("%s -> %s" % (x,y))
    G.add_edge(x,y)

#set default size and font for all nodes
for x in G:
    x.attr['fontname']='Arial bold'
    x.attr['fontsize']=16
    x.attr['color']='#ffffff'
    x.attr['style']='filled'
    x.attr['fillcolor']='#c4c4c4'

#change the shapes for selected nodes
change_attributes(G,parents,parent_shape,'shape')
change_attributes(G,children,child_shape,'shape')
G.write('graph.dot')

layout_programs = ['dot','neato','fdp']
for p in layout_programs:
    G.draw('graph_' + p + '.png', format='png',prog=p)
    print("Printed graph with %s." % p)


