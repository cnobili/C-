using System;
using System.Collections;
using Microsoft.SqlServer.Dts.Pipeline;
using Microsoft.SqlServer.Dts.Pipeline.Wrapper;
using Microsoft.SqlServer.Dts.Runtime.Wrapper;

namespace SSIS_CustomComponentUppercase
{
  [DtsPipelineComponent(DisplayName = "Uppercase", ComponentType = ComponentType.Transform)]
  public class Uppercase : PipelineComponent
  {
    ArrayList m_ColumnIndexList = new ArrayList();

    public override void ProvideComponentProperties()
    {
      base.ProvideComponentProperties();
      ComponentMetaData.InputCollection[0].Name = "Uppercase Input";
      ComponentMetaData.OutputCollection[0].Name = "Uppercase Output";
    }

    public override void PreExecute()
    {
      IDTSInput100 input = ComponentMetaData.InputCollection[0];
      IDTSInputColumnCollection100 inputColumns = input.InputColumnCollection;

      foreach (IDTSInputColumn100 column in inputColumns)
      {
        if (column.DataType == DataType.DT_STR || column.DataType == DataType.DT_WSTR)
        {
          m_ColumnIndexList.Add((int)BufferManager.FindColumnByLineageID(input.Buffer, column.LineageID));
        }
      }
    }

    public override void ProcessInput(int inputID, PipelineBuffer buffer)
    {
      while (buffer.NextRow())
      {
        foreach (int columnIndex in m_ColumnIndexList)
        {
          string str = buffer.GetString(columnIndex);
          buffer.SetString(columnIndex, str.ToUpper());
        }
      }
    }

  } // Class Uppercase

} // namespace SSIS_CustomComponentUppercase

