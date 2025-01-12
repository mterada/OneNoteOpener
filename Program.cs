using System;
using System.IO;
//using System.Collections.Generic;
//using System.Linq;
//using System.Threading.Tasks;
using System.Windows.Forms;
//using System.Runtime.InteropServices;
//using IWshRuntimeLibrary;
using System.Text.RegularExpressions;
using System.Text;

namespace OneNoteOpener
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Get URL from clipboard
            string url = Clipboard.GetText();

            // url is empty
            if (string.IsNullOrEmpty(url))
            {
                // message box
                MessageBox.Show("クリップボードにテキストがありません。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // url文字列から onenote: で始まる部分を取り出す
            int index = url.IndexOf("onenote:", StringComparison.Ordinal);
            if (index < 0)
            {
                // message box
                MessageBox.Show("OneNote URL がクリップボードにありません。 ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string onenoteUrl = url.Substring(index);
            string onenoteName = "";

            // onenoteUrl が "/" で終わっていたら ノートブック名を取り出す
            if (onenoteUrl.EndsWith("/")){
                // その文字を削除する
                string temp = onenoteUrl.Substring(0, onenoteUrl.Length - 1);
                // その１つ前の / から後ろを取り出す
                onenoteName = temp.Substring(temp.LastIndexOf("/", StringComparison.Ordinal) + 1);
            }
            // onenoteUrl が "&end"で終わっていたら セクションかページ
            else if (onenoteUrl.EndsWith("&end")){
                // "#section-id" の位置を取得
                int index2 = onenoteUrl.IndexOf(".one#section-id", StringComparison.Ordinal);
                int index3 = onenoteUrl.IndexOf("&section-id", StringComparison.Ordinal);

                if (index2 > 0){
                    // #の前にある / 以降の文字列を取り出す
                    string temp = onenoteUrl.Substring(0, index2) ;
                    onenoteName = temp.Substring(temp.LastIndexOf("/", StringComparison.Ordinal) + 1);
                }
                else if (index3 > 0){
                    // &の前にある / 以降の文字列を取り出す
                    string temp = onenoteUrl.Substring(0, index3) ;
                    onenoteName = temp.Substring(temp.LastIndexOf("#", StringComparison.Ordinal) + 1);
                }
                else {
                    // message box
                    MessageBox.Show("onenote: URL が不正です。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else {
                // message box
                MessageBox.Show("onenote: URL が不正です。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // onenoteName を安全なファイル名に変換して、ショートカット名にする
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string shortcutName = Regex.Replace(onenoteName, "[\\\\/:*?\"<>|]", "_") + ".url";
            // エスケープ文字を通常文字に戻す
            shortcutName = Uri.UnescapeDataString(shortcutName); 
            string shortcutPath = System.IO.Path.Combine(desktopPath, shortcutName );

            // Create a shortcut on the desktop
            using (StreamWriter writer = new StreamWriter(shortcutPath, false, Encoding.Default))
            {
                writer.WriteLine("[InternetShortcut]");
                writer.WriteLine("URL=" + onenoteUrl);
                writer.Close();
            }
            // 通知
            MessageBox.Show( "OneNoteショートカットを作成しました。\n\n" + onenoteUrl, shortcutName, MessageBoxButtons.OK );

            // クリップボードをクリアする
            Clipboard.Clear();

            return ;
            /*
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
            */
        }
    }
}
