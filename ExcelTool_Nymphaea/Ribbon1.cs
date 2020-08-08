using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExcelTool_Nymphaea
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button_Insert_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook xlWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet xlWorkSheet = xlWorkBook.ActiveSheet;

            String Extention = editBox_Extention.Text;
            String FileNameAddress = editBox_FileNameCol.Text;
            String FileInsertAddress = editBox_InsertCol.Text;
            String FilePath = editBox_FilePath.Text;

            DirectoryInfo directoryInfo = new DirectoryInfo(FilePath);

            //삽입할 이미지 파일 경로 유효성 체크
            if (!ValidateData.Path(directoryInfo))
            {
                editBox_FilePath.Text = string.Empty;
                return;
            }
            else if (string.IsNullOrEmpty(FilePath))
            {
                MessageBox.Show("파일 경로를 입력해주세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                return;
            }

            if (!ValidateData.CellAdress(xlWorkSheet, FileNameAddress))
            {
                editBox_FileNameCol.Text = string.Empty;

                return;
            }

            if (!ValidateData.CellAdress(xlWorkSheet, FileInsertAddress))
            {
                editBox_InsertCol.Text = string.Empty;

                return;
            }

            //확장자를 입력하지 않으면 자동으로 png 세팅
            if (string.IsNullOrEmpty(Extention))
                Extention = "png";

            int row = xlWorkSheet.Range[FileNameAddress].Row;

            for (; ; )
            {
                var fileName = xlWorkSheet.Cells[row, xlWorkSheet.Range[FileNameAddress].Column].Value;

                //더 이상 입력된 값이 없다면 종료
                if (fileName == null) break;

                //기존 이미지 지우기
                foreach (Excel.Shape image in xlWorkSheet.Shapes)
                {
                    if (image.TopLeftCell.Address == xlWorkSheet.Cells[row, xlWorkSheet.Range[FileInsertAddress].Column].Address)
                    {
                        image.Delete();
                    }
                }

                //새로운 이미지 삽입하기
                foreach (FileInfo fileInfo in directoryInfo.GetFiles($"{fileName}.{Extention}", SearchOption.AllDirectories))
                {
                    Excel.Range InsertRange = xlWorkSheet.Cells[row, xlWorkSheet.Range[FileInsertAddress].Column];

                    float left = (float)((double)InsertRange.Left);
                    float top = (float)((double)InsertRange.Top);
                    float width = (float)((double)xlWorkSheet.Cells[row, xlWorkSheet.Range[FileInsertAddress].Column + 1].Left) - (float)((double)InsertRange.Left);
                    float height = (float)((double)xlWorkSheet.Cells[row + 1, xlWorkSheet.Range[FileInsertAddress].Column].Top) - (float)((double)InsertRange.Top);

                    xlWorkSheet.Shapes.AddPicture($@"{FilePath}\{fileInfo.Name}", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, left, top, width, height);
                }

                row++;
            }

            MessageBox.Show("이미지 삽입 완료", "성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button_Delete_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook xlWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet xlWorkSheet = xlWorkBook.ActiveSheet;

            String FileNameAddress = editBox_FileNameCol.Text;
            String FileInsertAddress = editBox_InsertCol.Text;

            if (!ValidateData.CellAdress(xlWorkSheet, FileNameAddress))
            {
                editBox_FileNameCol.Text = string.Empty;
                return;
            }

            if (!ValidateData.CellAdress(xlWorkSheet, FileInsertAddress))
            {
                editBox_InsertCol.Text = string.Empty;
                return;
            }

            int row = xlWorkSheet.Range[FileNameAddress].Row;

            for (; ; )
            {
                var fileName = xlWorkSheet.Cells[row, xlWorkSheet.Range[FileNameAddress].Column].Value;

                if (fileName == null) break;

                foreach (Excel.Shape image in xlWorkSheet.Shapes)
                {
                    if (image.TopLeftCell.Address == xlWorkSheet.Cells[row, xlWorkSheet.Range[FileInsertAddress].Column].Address)
                    {
                        image.Delete();
                    }
                }
                row++;
            }

            if (!string.IsNullOrEmpty(editBox_FileNameCol.Text))
            {
                editBox_FileNameCol.Text = string.Empty;
            }

            if (!string.IsNullOrEmpty(editBox_InsertCol.Text))
            {
                editBox_InsertCol.Text = string.Empty;
            }

            MessageBox.Show("이미지 삭제 완료", "성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    public class LetterConverter
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Alphabet"></param>
        /// <returns></returns>
        public static int LetterToNumber(string letter)
        {
            /* 유효성 검사 throw new Exception();*/
            if (string.IsNullOrEmpty(letter)) throw new ArgumentNullException("Alphabet is null");

            letter = letter.ToUpperInvariant();

            int ASCIICode = 0;

            for (int i = 0; i < letter.Length; i++)
            {
                ASCIICode = Convert.ToInt32(letter[i]) - 64 + ASCIICode * 26;
            }

            return ASCIICode;
        }

        public static string NumberToLetter(int num)
        {
            //사용하지 않음
            return null;
        }
    }

    public class ValidateData
    {
        public static bool CellAdress(Excel.Worksheet xlws, String s)
        {
            String pattern = @"^[a-zA-Z]+[0-9]*$";
            String alphaStr = Regex.Replace(s, @"\d", "");
            String numStr = Regex.Replace(s, @"\D", "");

            if (string.IsNullOrEmpty(s)) //입력된 값이 존재하는 지 체크.
            {
                MessageBox.Show("셀 값을 입력해주세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            else if (!Regex.IsMatch(s, pattern)) //입력된 값이 영문자로 시작하고, 영문+숫자로 이루어져 있는 지 체크
            {
                MessageBox.Show("올바른 셀 값을 입력해주세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            else if (string.IsNullOrEmpty(alphaStr) || string.IsNullOrEmpty(numStr)) //영문, 숫자 모두 입력되어 있는 지 체크
            {
                MessageBox.Show("올바른 셀 값을 입력해주세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            int rowAddress = int.Parse(numStr);
            int colAddress = LetterConverter.LetterToNumber(alphaStr);

            if (rowAddress < 1 || rowAddress > xlws.Rows.Count || colAddress < 1 || colAddress > xlws.Columns.Count)
            {
                MessageBox.Show("해당 셀을 찾을 수 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                return false;
            }

            return true;
        }

        public static bool Path(DirectoryInfo di)
        {
            if (!di.Exists)
            {
                MessageBox.Show("지정된 경로가 올바르지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                return false;
            }

            return true;
        }
    }
}
