using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Text;
using Sample;

/// <summary>
/// UseADConnectSample の概要の説明です
/// </summary>
public class UseADConnectSample
{
    public UseADConnectSample()
    {
    }

    /// <summary>
    /// 検索のサンプル。
    /// </summary>
    public void doSearchSamplle()
    {
        ADConnectSample adConnSample = new Sample.ADConnectSample();
        List<Sample.ADAccountInfo> resultList = adConnSample.SearchByMail("@gmail.com");
        StringBuilder sb = new StringBuilder();
        foreach (Sample.ADAccountInfo info in resultList)
        {
            sb.Append(info.ToString());
        }
        Console.WriteLine(sb.ToString());
    }
}