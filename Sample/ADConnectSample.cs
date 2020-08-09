using System;
using System.Collections.Generic;
using System.DirectoryServices;

namespace Sample
{
    /// <summary>
    /// ActiveDirectory用接続クラスサンプル。ADと接続し、各種情報を取得します。
    /// </summary>
    public class ADConnectSample
    {
        // ADスキーマ上のアトリビュート名の定数群。
        // Ldap-Dispay-Name を検索時に使う。
        // サムアカウントネーム
        private string LDN_SAM_ACCOUNT_NAME = "sAMAccountName";
        //メールアドレス
        private string LDN_MAIL = "mail";
        // 表示名
        private string LDN_DISPLAY_NAME = "displayName";
        // 姓
        private string LDN_SN = "sn";
        // 名
        private string LDN_GIVEN_NAME = "givenName";

        /// <summary>
        /// 検索に使うディレクトリサーチャのリスト
        /// </summary>
        public List<DirectorySearcher> DsList { get; set; }

        /// <summary>
        /// デフォルトコンストラクタ。
        /// </summary>
        public ADConnectSample()
        {
            this.Initialize();
        }

        #region 内部処理
        /// <summary>
        /// 初期化処理を行います。
        /// </summary>
        private void Initialize()
        {
            // 以下を適宜変更する
            // {マシン名かIP：port}/{検索先のディレクトリをカンマ区切りで下層→上層に指定}
            string serverPath1 = "'11.22.33.44:5555/OU=SampleBaseDNDir,OU=moreDir1,DC=domaincntrl,DC=local";
            string serverPath2 = "'11.22.33.44:5555/OU=SampleBaseDNDir,OU=moreDir2,DC=domaincntrl,DC=local";
            string adUser = "domain\\username";
            string adPass = "password";

            // ディレクトリエントリはパスの数だけ指定する。
            string path1 = @"LDAP://" + serverPath1;
            DirectoryEntry de1 = new DirectoryEntry(path1, adUser, adPass);
            string path2 = @"LDAP://" + serverPath2;
            DirectoryEntry de2 = new DirectoryEntry(path2, adUser, adPass);

            List <DirectoryEntry> deList = new List<DirectoryEntry>();
            deList.Add(de1);
            deList.Add(de2);
            // ディレクトリサーチャーも複数メンバに持つようにしてみる。
            this.DsList = this.BuildDSList(deList);

            this.CheckIsConnected(deList);
        }

        /// <summary>
        /// ADとの接続確認を行います。
        /// <param name="deList">使用するディレクトリエントリのリスト</param>
        /// </summary>
        private void CheckIsConnected(List<DirectoryEntry> deList)
        {
            try
            {
                foreach (DirectoryEntry de in deList)
                {
                    Object obj = de.NativeObject;
                }
            }
            catch (Exception e)
            {
                throw new Exception("ADサーバとの接続でエラーが発生しました。", e);
            }
        }


        /// <summary>
        /// 検索結果１件からAD情報へ変換します。
        /// </summary>
        /// <param name="sResult">検索結果のオブジェクト</param>
        /// <returns>ADAccountInfoインスタンス</returns>
        private ADAccountInfo ConvSearchResult(SearchResult sResult)
        {
            if (sResult == null)
            {
                return null;
            }

            ADAccountInfo adInfo = new ADAccountInfo();
            // 該当する属性があったらセットしていく。
            if (sResult.Properties[LDN_SAM_ACCOUNT_NAME].Count > 0)
            {
                adInfo.SamAccountName = (string)sResult.Properties[LDN_SAM_ACCOUNT_NAME][0];
            }
            if (sResult.Properties[LDN_MAIL].Count > 0)
            {
                adInfo.Mail = (string)sResult.Properties[LDN_MAIL][0];
            }
            if (sResult.Properties[LDN_DISPLAY_NAME].Count > 0)
            {
                adInfo.DisplayName = (string)sResult.Properties[LDN_DISPLAY_NAME][0];
            }
            if (sResult.Properties[LDN_SN].Count > 0)
            {
                adInfo.Sn = (string)sResult.Properties[LDN_SN][0];
            }
            if (sResult.Properties[LDN_GIVEN_NAME].Count > 0)
            {
                adInfo.GivenName = (string)sResult.Properties[LDN_GIVEN_NAME][0];
            }

            return adInfo;
        }

        /// <summary>
        /// 検索結果のコレクションをリストに変換します。
        /// </summary>
        /// <param name="resultCollection">ディレクトリサービス用の検索結果コレクション</param>
        /// <returns>ADAccountInfoインスタンスのリスト(結果なしの場合は0件リスト)</returns>
        private List<ADAccountInfo> ConvSearchResultCollection(SearchResultCollection resultCollection)
        {
            List<ADAccountInfo> adInfoList = new List<ADAccountInfo>();
            if (resultCollection == null || resultCollection.Count == 0)
            {
                return adInfoList;
            }
            foreach (SearchResult sResult in resultCollection)
            {
                adInfoList.Add(this.ConvSearchResult(sResult));
            }
            return adInfoList;
        }


        /// <summary>
        /// 内部的なディレクトリサーチャーのリストを組み立てます。
        /// </summary>
        /// <param name="deList">ディレクトリエントリのリスト</param>
        /// <returns>ディレクトリサーチャーのリスト</returns>
        private List<DirectorySearcher> BuildDSList(List<DirectoryEntry> deList)
        {
            List<DirectorySearcher> dsList = new List<DirectorySearcher>();
            // LDAP検索オブジェクトを作成
            foreach (DirectoryEntry de in deList)
            {
                DirectorySearcher ds = new DirectorySearcher(de);
                dsList.Add(ds);
            }
            return dsList;
        }
        #endregion

        #region 外部API的なpubicメソッド
        /// <summary>
        /// メールアドレスの部分一致でADアカウントを検索します。
        /// </summary>
        /// <param name="mail">メールアドレスの検索文字列</param>
        /// <returns>ADInfoのリスト(存在しなかったら0件)</returns>
        public List<ADAccountInfo> SearchByMail(string mail)
        {
            // フィルターを設定
            string filter = "(mail=*" + mail + "*)";
            List<ADAccountInfo> adInfoList = new List<ADAccountInfo>();
            foreach (DirectorySearcher ds in this.DsList)
            {
                ds.Filter = filter;
                // 個別に取得する属性を設定
                ds.PropertiesToLoad.Add(LDN_SAM_ACCOUNT_NAME);
                ds.PropertiesToLoad.Add(LDN_DISPLAY_NAME);
                ds.PropertiesToLoad.Add(LDN_MAIL);
                ds.PropertiesToLoad.Add(LDN_SN);
                ds.PropertiesToLoad.Add(LDN_GIVEN_NAME);

                // 検索を実行
                SearchResultCollection resultCollection = null;
                try
                {
                    resultCollection = ds.FindAll();
                }
                catch (Exception e)
                {
                    continue; // ここでは握り潰し
                }
                adInfoList.AddRange(this.ConvSearchResultCollection(resultCollection));
            }
            return adInfoList;
        }
        #endregion

    }

    /// <summary>
    /// Active Directory情報クラス。ADから取得したアカウント１件分の情報を保持します。
    /// </summary>
    public class ADAccountInfo
    {
        /// <summary>
        /// 名前
        /// </summary>
        public string SamAccountName { get; set; }

        /// <summary>
        /// メールアドレス
        /// </summary>
        public string Mail { get; set; }

        /// <summary>
        /// 表示名
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// 姓
        /// </summary>
        public string Sn { get; set; }

        /// <summary>
        /// 名
        /// </summary>
        public string GivenName { get; set; }


        /// <summary>
        /// コンストラクタ。
        /// </summary>
        public ADAccountInfo()
        {
        }

        /// <summary>
        /// 内容を文字列で表示します。確認用です。
        /// </summary>
        /// <returns>インスタンスの内容を表す文字列</returns>
        public string ToString()
        {
            string str = " samAccountName: " + this.SamAccountName
                + " mail: " + this.Mail
                + " displayName: " + this.DisplayName
                + " surName: " + this.Sn
                + " givenName: " + this.GivenName
                + System.Environment.NewLine
                ;
            return str;
        }
    }
}