using Microsoft.VisualBasic;

internal class Convertion
{
    private object mvarvar; // local copy
    private string mvarFormato; // local copy
    private object mvarfileName; // local copy
    private short mvarLineNum; // local copy
    private short mvarInicia; // local copy
    private short mvarLeng; // local copy
    private string mvarCondition; // local copy
    private short mvarcol1; // local copy
    private short mvarcol2; // local copy
                            // local variable(s) to hold property value(s)
    private object mvarDescription; // local copy

    public string Description
    {
        get
        {
            return (string)mvarDescription;
        }
        set
        {
            mvarDescription = value;
        }
    }
    public short col2
    {
        set
        {
            mvarcol2 = value;
        }
    }
    public short col1
    {
        set
        {
            mvarcol1 = value;
        }
    }

    public string Condition
    {
        set
        {
            mvarCondition = value;
        }
    }
    public short Leng
    {
        set
        {
            mvarLeng = value;
        }
    }
    public short Inicia
    {
        set
        {
            mvarInicia = value;
        }
    }
    public short LineNum
    {
        set
        {
            mvarLineNum = value;
        }
    }

    public string filename
    {
        set
        {
           mvarfileName = value;
        }
    }

    public string convert(ref string mvarvar, ref string mvarFormato)
    {
        string newval;
        string formated;
        double mvarvarnum;
        int Leng, lenfor;
        int i;

        string conversion = "";
        mvarvarnum = double.Parse(mvarvar);
        formated = string.Format("{0:" + mvarFormato + "}", mvarvarnum);
        // formated = String.Format(mvarvar, mvarFormato)
        Leng = formated.Length;
        lenfor = mvarFormato.Length;

        for (i = Leng + 1; i <= lenfor; i++) 
            conversion = conversion + " ";

        if (mvarFormato[0] == '#')
            conversion = conversion + formated;
        else
            conversion = formated + conversion;

        return conversion;
    }

    public string value()
    {
        /*
        object i;
        object mypos;
        object COND1;
        object z;
        object fs;
        string[] values;

        fs = Interaction.CreateObject("Scripting.FileSystemObject");
        z = fs.OpenTextFile(mvarfileName);

        if ((Strings.Trim(mvarCondition) != ""))
        {
            COND1 = z.ReadLine;
            mypos = Strings.InStr(1, COND1, mvarCondition);
            while (mypos == 0)
            {
                COND1 = z.ReadLine;
                mypos = Strings.InStr(1, COND1, mvarCondition);
            }
            value = Strings.Mid(COND1, mvarInicia, 16);
        }
        else
        {
            for (i = 1; i <= mvarLineNum - 1; i++)
                z.ReadLine();
            value = Strings.Mid(z.ReadLine, mvarInicia, mvarLeng);
            if (value.Contains("|"))
            {
                values = Strings.Split(value, "|");
                value = values[0];
            }
        }

        return;
        goError:
        ;
        if (Information.Err.Number == 53)
            Interaction.MsgBox(Information.Err.Description + mvarfileName);
        else
            Interaction.MsgBox(Information.Err.Description);
            */
        return "";
    }
}
