import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
import System
from System.Windows.Forms import *
from System.Drawing import *

def multiplecombo(title,ltexts,lists):
    form1 = Form()
    form1.Text = title
    labels=[]
    combos=[]
    
    x=0
    for a in ltexts:
        label1 = Label()
        label1.Text=str(a)
        label1.Location = Point (20,10+x)
        label1.Size = Size(400,20)
        labels.append(label1)
        x=x+50
    
    y=0
    for a in lists:
        combo = ComboBox()
        for b in a:
    	    combo.Items.Add(b)
        combo.Name = "Combo"
        combo.Location = Point (20,30+y)
        combo.Size = Size(400,20)
        combos.append(combo)
        y=y+50
   
    button1 = Button()
    button1.Text = "Yes"
    button1.DialogResult = DialogResult.Yes
    button1.Location = Point(20,y+40)
       
    form1.HelpButton = False
    form1.FormBorderStyle = FormBorderStyle.FixedToolWindow
    form1.MaximizeBox = False
    form1.MinimizeBox = False
    form1.AcceptButton = button1
    form1.StartPosition = FormStartPosition.CenterScreen

    form1.Controls.Add(button1)
    for a in combos:
        form1.Controls.Add(a)
    for a in labels:
        form1.Controls.Add(a)
    
    form1.Size = Size(460,y+120)
    form1.ShowDialog()
    results=[]
    listofvalues = form1.Controls.Find("Combo",True)

    for a in listofvalues:
        results.append(a.SelectedItem)
 

    return results
