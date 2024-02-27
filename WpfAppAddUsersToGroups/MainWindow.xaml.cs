using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.DirectoryServices;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Xml.Linq;
using WpfAppAddUsersToGroups.Extensions;


namespace WpfAppAddUsersToGroups
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        enum ResultAddToGroup
        {
            NotAdd,
            AlreadyAdded,
            Add
        }


        enum ResultUserInGroup
        {
            InGroup,
            NotInGroup,
            NotFind
        }

        enum ResultFind
        {
            Find,
            NotFind,
            FindMore1
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        record FindResult(string findStr, ResultFind resultFind);

        record UserFindResult(string findStr, ResultFind resultFind, string login, string userDn) : FindResult(findStr, resultFind);
        record GroupFindResult(string findStr, ResultFind resultFind, string groupCn, string groupAdspath, string[] members) : FindResult(findStr, resultFind);


        record AddGroupResult(UserFindResult[] resultFindUsers, GroupFindResult[] resultFindGroups, ResultAddToGroup[,] resultAddToGroup);

        record UserInGroupResult(UserFindResult[] resultFindUsers, GroupFindResult[] resultFindGroups, ResultUserInGroup[,] resultUserInGroups);
        
        private void ButtonSeeUsersInGroup_Click(object sender, RoutedEventArgs e)
        {
            string[] findNames = TextBoxNamesInput.Text.Split("\r\n")
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .ToArray();

            string[] findGroups = TextBoxGroupsInput.Text.Split("\r\n")
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .ToArray();

            TextBoxNamesInput.Text = string.Join("\r\n", findNames);
            TextBoxGroupsInput.Text = string.Join("\r\n", findGroups);


            UserInGroupResult resultuserInGroup = UserInGroup(findNames, findGroups);

            SaveToExel(resultuserInGroup);
        }


        private void ButtonAddToGroup_Click(object sender, RoutedEventArgs e)
        {
            string[] findNames = TextBoxNamesInput.Text.Split("\r\n")
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .ToArray();

            string[] findGroups = TextBoxGroupsInput.Text.Split("\r\n")
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .ToArray();

            TextBoxNamesInput.Text = string.Join("\r\n", findNames);
            TextBoxGroupsInput.Text = string.Join("\r\n", findGroups);


            AddGroupResult resultAddGroup = AddToGroup(findNames, findGroups);

            SaveToExel(resultAddGroup);
        }


        private UserInGroupResult UserInGroup(string[] findNames, string[] findGroups)
        {
            DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://DC=office,DC=crocusgroup,DC=ru");

            UserFindResult[] resultFindUsers = FindUsers(directoryEntry, findNames);
            GroupFindResult[] resultFindGroups = FindGroups(directoryEntry, findGroups);

            ResultUserInGroup[,] resultUserInGroup = new ResultUserInGroup[findGroups.Length, findNames.Length];


            for (int i = 0; i < resultFindGroups.Length; i++)
            {
                DirectoryEntry directoryEntryGroup = resultFindGroups[i].resultFind == ResultFind.Find ?
                    new DirectoryEntry(resultFindGroups[i].groupAdspath) : null;

                for (int j = 0; j < resultFindUsers.Length; j++)
                {
                    resultUserInGroup[i, j] = (resultFindGroups[i].resultFind, resultFindUsers[j].resultFind) switch
                    {
                        (ResultFind.Find, ResultFind.Find) => resultFindGroups[i].members.Contains(resultFindUsers[j].login) ? ResultUserInGroup.InGroup : ResultUserInGroup.NotInGroup,
                        _ => ResultUserInGroup.NotFind
                    };

                    /*
                    if (resultFindGroups[i].resultFind == ResultFind.Find && resultFindUsers[j].resultFind == ResultFind.Find)
                    {
                        if (resultFindGroups[i].members.Contains(resultFindUsers[j].login))
                            resultUserInGroup[i, j] = ResultUserInGroup.InGroup;
                        else
                            resultUserInGroup[i, j] = ResultUserInGroup.NotInGroup;
                    }
                    else
                        resultUserInGroup[i, j] = ResultUserInGroup.NotFind;
                    */
                }

                if (resultFindGroups[i].resultFind == ResultFind.Find)
                    directoryEntryGroup.CommitChanges();
            }

            return new UserInGroupResult(resultFindUsers, resultFindGroups, resultUserInGroup);
        }



        private AddGroupResult AddToGroup(string[] findNames, string[] findGroups)
        {
            DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://DC=office,DC=crocusgroup,DC=ru");

            UserFindResult[] resultFindUsers = FindUsers(directoryEntry, findNames);
            GroupFindResult[] resultFindGroups = FindGroups(directoryEntry, findGroups);


            ResultAddToGroup[,] resultAddToGroups = new ResultAddToGroup[findGroups.Length, findNames.Length];

            
            for (int i = 0; i < resultFindGroups.Length; i++)
            {
                DirectoryEntry directoryEntryGroup = resultFindGroups[i].resultFind == ResultFind.Find ?
                    new DirectoryEntry(resultFindGroups[i].groupAdspath) : null;

                for (int j = 0; j < resultFindUsers.Length; j++)
                {
                    if (resultFindGroups[i].resultFind == ResultFind.Find && resultFindUsers[j].resultFind == ResultFind.Find)
                    {
                        if (resultFindGroups[i].members.Contains(resultFindUsers[j].login))
                            resultAddToGroups[i, j] = ResultAddToGroup.AlreadyAdded;
                        else
                        {
                            directoryEntryGroup.Properties["member"].Add(resultFindUsers[j].userDn);
                            resultAddToGroups[i, j] = ResultAddToGroup.Add;
                        }
                    }
                    else
                        resultAddToGroups[i, j] = ResultAddToGroup.NotAdd;
                }

                if (resultFindGroups[i].resultFind == ResultFind.Find)
                    directoryEntryGroup.CommitChanges();
            }

            return new AddGroupResult(resultFindUsers, resultFindGroups, resultAddToGroups);
        }


        private UserFindResult[] FindUsers(DirectoryEntry directoryEntry, string[] findNames) =>
            findNames.Select(findName =>
            {
                DirectorySearcher dirSearchUser = new DirectorySearcher(directoryEntry, $"(&(objectClass=user)(objectCategory=person)(anr={findName}))");

                SearchResultCollection searchResColl = dirSearchUser.FindAll();

                return searchResColl.Count switch
                {
                    0 => new UserFindResult(findName, ResultFind.NotFind, "Not find.", string.Empty),
                    1 => new UserFindResult(findName, ResultFind.Find, searchResColl[0].GetProp("sAMAccountName"), searchResColl[0].GetProp("distinguishedname")),
                    _ => new UserFindResult(findName, ResultFind.FindMore1, "Find more1.", string.Empty)
                };
            }).ToArray();

        private GroupFindResult[] FindGroups(DirectoryEntry directoryEntry, string[] findGroups) =>
            findGroups.Select(findGroup =>
            {
                DirectorySearcher dirSearchGroup = new DirectorySearcher(directoryEntry, $"(&(objectClass=group)(objectCategory=group)(cn={findGroup}))");
                SearchResultCollection searchResColl = dirSearchGroup.FindAll();

                return searchResColl.Count switch
                {
                    0 => new GroupFindResult(findGroup, ResultFind.NotFind, "Not find.", string.Empty, Array.Empty<string>()),
                    1 => new GroupFindResult(findGroup, ResultFind.Find, searchResColl[0].GetProp("cn"), searchResColl[0].GetProp("adspath"), GetMembersGroup(directoryEntry, searchResColl[0].GetProp("distinguishedname"))),
                    _ => new GroupFindResult(findGroup, ResultFind.FindMore1, "Not find.", string.Empty, Array.Empty<string>())
                };
            }).ToArray();

        private string[] GetMembersGroup(DirectoryEntry directoryEntry, string dnGroup)
        {
            DirectorySearcher dirSearchMemberGroup = new DirectorySearcher(directoryEntry, $"(&(objectClass=user)(objectCategory=user)(memberOf={dnGroup}))");

            SearchResultCollection searchResCollUsersGroup = dirSearchMemberGroup.FindAll();

            return searchResCollUsersGroup.Cast<SearchResult>().Select(user => user.GetProp("sAMAccountName")).ToArray();
        }


        private void SaveToExel(AddGroupResult data)
        {
            Dictionary<ResultFind, string> resultFindStatusToString = new Dictionary<ResultFind, string>()
            {
                [ResultFind.Find] = "Найден{0}.",
                [ResultFind.NotFind] = "Не найден{0}.",
                [ResultFind.FindMore1] = "Найден{0} более 1."
            };

            Dictionary<ResultAddToGroup, string> resultAddStatusToString = new Dictionary<ResultAddToGroup, string>()
            {
                [ResultAddToGroup.Add] = "Добавлен.",
                [ResultAddToGroup.AlreadyAdded] = "Уже в группе.",
                [ResultAddToGroup.NotAdd] = "Не добавлен.",
            };

            string GetStrFindStatUser(ResultFind resultFind) =>
                string.Format(resultFindStatusToString[resultFind], "");

            string GetStrFindStatGroup(ResultFind resultFind) =>
                string.Format(resultFindStatusToString[resultFind], "а");

            string GetColLet(int col) =>
                ExcelCellAddress.GetColumnLetter(col);

            ExcelPackage excelPackage = new ExcelPackage();

            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Result add to groups.");

            int startRow = 2;
            int startCol = 2;

            int dataOffsetRow = 3;
            int dataOffsetColl = 3;


            #region headerAndF
            worksheet.Cells[startRow + 2, startCol].Value = "Login/Name/Email";
            worksheet.Cells[startRow + 2, startCol + 1].Value = "Status Find";

            worksheet.Cells[startRow, startCol + 2].Value = "Group Find:";
            worksheet.Cells[startRow + 1, startCol + 2].Value = "Status Find:";

            var cellLoginGroup = worksheet.Cells[startRow + 2, startCol + 2];
            cellLoginGroup.IsRichText = true;
            var loginText = cellLoginGroup.RichText.Add("Login");
            cellLoginGroup.RichText.Add(" ");
            var groupText = cellLoginGroup.RichText.Add("Group");
            //worksheet.Cells[startRow + 1, startCol + 1].Value = "Login Group";
            loginText.Size = 16;
            groupText.Size = 16;
            loginText.VerticalAlign = ExcelVerticalAlignmentFont.Subscript;
            groupText.VerticalAlign = ExcelVerticalAlignmentFont.Superscript;
            cellLoginGroup.Style.Border.Diagonal.Style = ExcelBorderStyle.Thin;
            cellLoginGroup.Style.Border.DiagonalDown = true;

            #region border
            //var cellsInnerBorder = worksheet.Cells[$"{GetColLet(startCol)}{startRow + 2}:{GetColLet(startCol + 2 + data.resultFindGroups.Length)}{startRow + 2 + data.resultFindUsers.Length},{GetColLet(startCol + 2)}{startRow}:{GetColLet(startCol + 2 + data.resultFindGroups.Length)}{startRow + 1}"];
            var cellsInnerBorder = worksheet.Cells[startRow, startCol, startRow + 2 + data.resultFindUsers.Length, startCol + 2 + data.resultFindGroups.Length];

            cellsInnerBorder.Style.Border.Top.Style = cellsInnerBorder.Style.Border.Bottom.Style =
                cellsInnerBorder.Style.Border.Left.Style = cellsInnerBorder.Style.Border.Right.Style =
                ExcelBorderStyle.Thin;


            #region aroundBorder
            worksheet.Cells[startRow, startCol, startRow + 2 + data.resultFindUsers.Length, startCol + 2 + data.resultFindGroups.Length]
                .Style.Border.BorderAround(ExcelBorderStyle.Medium);

            var cellsBorderMinus = worksheet.Cells[startRow, startCol, startRow + 1, startCol + 1];
            cellsBorderMinus.Style.Border.Top.Style = cellsBorderMinus.Style.Border.Bottom.Style =
                cellsBorderMinus.Style.Border.Left.Style = cellsBorderMinus.Style.Border.Right.Style =
                ExcelBorderStyle.None;

            worksheet.Cells[startRow + 1, startCol, startRow + 1, startCol + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            worksheet.Cells[startRow, startCol + 1, startRow + 1, startCol + 1].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            #endregion
            #endregion

            var cellsHeader = worksheet.Cells[startRow, startCol, startRow + 2, startCol + 2 + data.resultFindGroups.Length];
            cellsHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cellsHeader.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            

            var cellsUsers = worksheet.Cells[startRow + 3, startCol, startRow + 3 + data.resultFindUsers.Length, startCol + 2];
            string condFormUserFormula = $"=${GetColLet(startCol + 1)}{startRow + 3}=\"{{0}}\"";

            var conditionalFormatUserFind = worksheet.ConditionalFormatting.AddExpression(cellsUsers);
            conditionalFormatUserFind.Formula = string.Format(condFormUserFormula, GetStrFindStatUser(ResultFind.Find));
            conditionalFormatUserFind.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatUserFind.Style.Fill.BackgroundColor.Color = Color.Green;

            var conditionalFormatUserNotFind = worksheet.ConditionalFormatting.AddExpression(cellsUsers);
            conditionalFormatUserNotFind.Formula = string.Format(condFormUserFormula, GetStrFindStatUser(ResultFind.NotFind));
            conditionalFormatUserNotFind.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatUserNotFind.Style.Fill.BackgroundColor.Color = Color.Red;

            var conditionalFormatUserFindMore1 = worksheet.ConditionalFormatting.AddExpression(cellsUsers);
            conditionalFormatUserFindMore1.Formula = string.Format(condFormUserFormula, GetStrFindStatUser(ResultFind.FindMore1));
            conditionalFormatUserFindMore1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatUserFindMore1.Style.Fill.BackgroundColor.Color = Color.Yellow;


            var cellsGroups = worksheet.Cells[startRow, startCol + 3, startRow + 2, startCol + 3 + data.resultFindGroups.Length];
            string condFormGroupFormula = $"={GetColLet(startCol + 3)}${startRow + 1}=\"{{0}}\"";

            var conditionalFormatGroupFind = worksheet.ConditionalFormatting.AddExpression(cellsGroups);
            conditionalFormatGroupFind.Formula = string.Format(condFormGroupFormula, GetStrFindStatGroup(ResultFind.Find));
            conditionalFormatGroupFind.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatGroupFind.Style.Fill.BackgroundColor.Color = Color.Green;

            var conditionalFormatGroupNotFind = worksheet.ConditionalFormatting.AddExpression(cellsGroups);
            conditionalFormatGroupNotFind.Formula = string.Format(condFormGroupFormula, GetStrFindStatGroup(ResultFind.NotFind));
            conditionalFormatGroupNotFind.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatGroupNotFind.Style.Fill.BackgroundColor.Color = Color.Red;

            var conditionalFormatGroupFindMore1 = worksheet.ConditionalFormatting.AddExpression(cellsGroups);
            conditionalFormatGroupFindMore1.Formula = string.Format(condFormGroupFormula, GetStrFindStatGroup(ResultFind.FindMore1));
            conditionalFormatGroupFindMore1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatGroupFindMore1.Style.Fill.BackgroundColor.Color = Color.Yellow;


            var cellsData = worksheet.Cells[startRow + dataOffsetRow, startCol + dataOffsetColl, startRow + dataOffsetRow + data.resultFindUsers.Length, startCol + dataOffsetColl + data.resultFindGroups.Length];

            var conditionalFormatDataAdd = worksheet.ConditionalFormatting.AddEqual(cellsData);
            conditionalFormatDataAdd.Formula = $"\"{resultAddStatusToString[ResultAddToGroup.Add]}\"";
            conditionalFormatDataAdd.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatDataAdd.Style.Fill.BackgroundColor.Color = Color.Green;

            var conditionalFormatDataAlreadyAdded = worksheet.ConditionalFormatting.AddEqual(cellsData);
            conditionalFormatDataAlreadyAdded.Formula = $"\"{resultAddStatusToString[ResultAddToGroup.AlreadyAdded]}\"";
            conditionalFormatDataAlreadyAdded.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatDataAlreadyAdded.Style.Fill.BackgroundColor.Color = Color.LightGreen;

            var conditionalFormatDataNotAdd = worksheet.ConditionalFormatting.AddEqual(cellsData);
            conditionalFormatDataNotAdd.Formula = $"\"{resultAddStatusToString[ResultAddToGroup.NotAdd]}\"";
            conditionalFormatDataNotAdd.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatDataNotAdd.Style.Fill.BackgroundColor.Color = Color.Red;
            
            #endregion

            for (int i = 0; i < data.resultFindGroups.Length; i++)
            {
                worksheet.Cells[startRow, startCol + 3 + i].Value = data.resultFindGroups[i].findStr;
                worksheet.Cells[startRow + 1, startCol + 3 + i].Value = GetStrFindStatGroup(data.resultFindGroups[i].resultFind);
                worksheet.Cells[startRow + 2, startCol + 3 + i].Value = data.resultFindGroups[i].groupCn;
            }

            for (int i = 0; i < data.resultFindUsers.Length; i++)
            {
                worksheet.Cells[startRow + 3 + i, startCol].Value = data.resultFindUsers[i].findStr;
                worksheet.Cells[startRow + 3 + i, startCol + 1].Value = GetStrFindStatUser(data.resultFindUsers[i].resultFind);
                worksheet.Cells[startRow + 3 + i, startCol + 2].Value = data.resultFindUsers[i].login;
            }

            for (int i = 0; i < data.resultAddToGroup.GetLength(0); i++)
            {
                for (int j = 0; j < data.resultAddToGroup.GetLength(1); j++)
                {
                    worksheet.Cells[startRow + dataOffsetRow + j, startCol + dataOffsetColl + i].Value = resultAddStatusToString[data.resultAddToGroup[i, j]];
                }
            }


            worksheet.Cells.AutoFitColumns();

            string path = $"{Path.GetTempFileName()}.xlsx";

            excelPackage.SaveAs(new FileInfo(path));

            Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE", path);
        }

        private void SaveToExel(UserInGroupResult data)
        {
            Dictionary<ResultFind, string> resultFindStatusToString = new Dictionary<ResultFind, string>()
            {
                [ResultFind.Find] = "Найден{0}.",
                [ResultFind.NotFind] = "Не найден{0}.",
                [ResultFind.FindMore1] = "Найден{0} более 1."
            };

            Dictionary<ResultUserInGroup, string> resultUserInGroupStatusToString = new Dictionary<ResultUserInGroup, string>()
            {
                [ResultUserInGroup.InGroup] = "В группе.",
                [ResultUserInGroup.NotInGroup] = "Не в группе.",
                [ResultUserInGroup.NotFind] = "не найден пользователь или группа.",
            };

            string GetStrFindStatUser(ResultFind resultFind) =>
                string.Format(resultFindStatusToString[resultFind], "");

            string GetStrFindStatGroup(ResultFind resultFind) =>
                string.Format(resultFindStatusToString[resultFind], "а");

            string GetColLet(int col) =>
                ExcelCellAddress.GetColumnLetter(col);

            ExcelPackage excelPackage = new ExcelPackage();

            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Result add to groups.");

            int startRow = 2;
            int startCol = 2;

            int dataOffsetRow = 3;
            int dataOffsetColl = 3;


            #region headerAndF
            worksheet.Cells[startRow + 2, startCol].Value = "Login/Name/Email";
            worksheet.Cells[startRow + 2, startCol + 1].Value = "Status Find";

            worksheet.Cells[startRow, startCol + 2].Value = "Group Find:";
            worksheet.Cells[startRow + 1, startCol + 2].Value = "Status Find:";

            var cellLoginGroup = worksheet.Cells[startRow + 2, startCol + 2];
            cellLoginGroup.IsRichText = true;
            var loginText = cellLoginGroup.RichText.Add("Login");
            cellLoginGroup.RichText.Add(" ");
            var groupText = cellLoginGroup.RichText.Add("Group");
            //worksheet.Cells[startRow + 1, startCol + 1].Value = "Login Group";
            loginText.Size = 16;
            groupText.Size = 16;
            loginText.VerticalAlign = ExcelVerticalAlignmentFont.Subscript;
            groupText.VerticalAlign = ExcelVerticalAlignmentFont.Superscript;
            cellLoginGroup.Style.Border.Diagonal.Style = ExcelBorderStyle.Thin;
            cellLoginGroup.Style.Border.DiagonalDown = true;

            #region border
            //var cellsInnerBorder = worksheet.Cells[$"{GetColLet(startCol)}{startRow + 2}:{GetColLet(startCol + 2 + data.resultFindGroups.Length)}{startRow + 2 + data.resultFindUsers.Length},{GetColLet(startCol + 2)}{startRow}:{GetColLet(startCol + 2 + data.resultFindGroups.Length)}{startRow + 1}"];
            var cellsInnerBorder = worksheet.Cells[startRow, startCol, startRow + 2 + data.resultFindUsers.Length, startCol + 2 + data.resultFindGroups.Length];

            cellsInnerBorder.Style.Border.Top.Style = cellsInnerBorder.Style.Border.Bottom.Style =
                cellsInnerBorder.Style.Border.Left.Style = cellsInnerBorder.Style.Border.Right.Style =
                ExcelBorderStyle.Thin;


            #region aroundBorder
            worksheet.Cells[startRow, startCol, startRow + 2 + data.resultFindUsers.Length, startCol + 2 + data.resultFindGroups.Length]
                .Style.Border.BorderAround(ExcelBorderStyle.Medium);

            var cellsBorderMinus = worksheet.Cells[startRow, startCol, startRow + 1, startCol + 1];
            cellsBorderMinus.Style.Border.Top.Style = cellsBorderMinus.Style.Border.Bottom.Style =
                cellsBorderMinus.Style.Border.Left.Style = cellsBorderMinus.Style.Border.Right.Style =
                ExcelBorderStyle.None;

            worksheet.Cells[startRow + 1, startCol, startRow + 1, startCol + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            worksheet.Cells[startRow, startCol + 1, startRow + 1, startCol + 1].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            #endregion
            #endregion

            var cellsHeader = worksheet.Cells[startRow, startCol, startRow + 2, startCol + 2 + data.resultFindGroups.Length];
            cellsHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cellsHeader.Style.VerticalAlignment = ExcelVerticalAlignment.Center;


            var cellsUsers = worksheet.Cells[startRow + 3, startCol, startRow + 3 + data.resultFindUsers.Length, startCol + 2];
            string condFormUserFormula = $"=${GetColLet(startCol + 1)}{startRow + 3}=\"{{0}}\"";

            var conditionalFormatUserFind = worksheet.ConditionalFormatting.AddExpression(cellsUsers);
            conditionalFormatUserFind.Formula = string.Format(condFormUserFormula, GetStrFindStatUser(ResultFind.Find));
            conditionalFormatUserFind.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatUserFind.Style.Fill.BackgroundColor.Color = Color.Green;

            var conditionalFormatUserNotFind = worksheet.ConditionalFormatting.AddExpression(cellsUsers);
            conditionalFormatUserNotFind.Formula = string.Format(condFormUserFormula, GetStrFindStatUser(ResultFind.NotFind));
            conditionalFormatUserNotFind.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatUserNotFind.Style.Fill.BackgroundColor.Color = Color.Red;

            var conditionalFormatUserFindMore1 = worksheet.ConditionalFormatting.AddExpression(cellsUsers);
            conditionalFormatUserFindMore1.Formula = string.Format(condFormUserFormula, GetStrFindStatUser(ResultFind.FindMore1));
            conditionalFormatUserFindMore1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatUserFindMore1.Style.Fill.BackgroundColor.Color = Color.Yellow;


            var cellsGroups = worksheet.Cells[startRow, startCol + 3, startRow + 2, startCol + 3 + data.resultFindGroups.Length];
            string condFormGroupFormula = $"={GetColLet(startCol + 3)}${startRow + 1}=\"{{0}}\"";

            var conditionalFormatGroupFind = worksheet.ConditionalFormatting.AddExpression(cellsGroups);
            conditionalFormatGroupFind.Formula = string.Format(condFormGroupFormula, GetStrFindStatGroup(ResultFind.Find));
            conditionalFormatGroupFind.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatGroupFind.Style.Fill.BackgroundColor.Color = Color.Green;

            var conditionalFormatGroupNotFind = worksheet.ConditionalFormatting.AddExpression(cellsGroups);
            conditionalFormatGroupNotFind.Formula = string.Format(condFormGroupFormula, GetStrFindStatGroup(ResultFind.NotFind));
            conditionalFormatGroupNotFind.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatGroupNotFind.Style.Fill.BackgroundColor.Color = Color.Red;

            var conditionalFormatGroupFindMore1 = worksheet.ConditionalFormatting.AddExpression(cellsGroups);
            conditionalFormatGroupFindMore1.Formula = string.Format(condFormGroupFormula, GetStrFindStatGroup(ResultFind.FindMore1));
            conditionalFormatGroupFindMore1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatGroupFindMore1.Style.Fill.BackgroundColor.Color = Color.Yellow;


            var cellsData = worksheet.Cells[startRow + dataOffsetRow, startCol + dataOffsetColl, startRow + dataOffsetRow + data.resultFindUsers.Length, startCol + dataOffsetColl + data.resultFindGroups.Length];

            var conditionalFormatDataAdd = worksheet.ConditionalFormatting.AddEqual(cellsData);
            conditionalFormatDataAdd.Formula = $"\"{resultUserInGroupStatusToString[ResultUserInGroup.InGroup]}\"";
            conditionalFormatDataAdd.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatDataAdd.Style.Fill.BackgroundColor.Color = Color.Green;

            var conditionalFormatDataAlreadyAdded = worksheet.ConditionalFormatting.AddEqual(cellsData);
            conditionalFormatDataAlreadyAdded.Formula = $"\"{resultUserInGroupStatusToString[ResultUserInGroup.NotInGroup]}\"";
            conditionalFormatDataAlreadyAdded.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatDataAlreadyAdded.Style.Fill.BackgroundColor.Color = Color.LightBlue;

            var conditionalFormatDataNotAdd = worksheet.ConditionalFormatting.AddEqual(cellsData);
            conditionalFormatDataNotAdd.Formula = $"\"{resultUserInGroupStatusToString[ResultUserInGroup.NotFind]}\"";
            conditionalFormatDataNotAdd.Style.Fill.PatternType = ExcelFillStyle.Solid;
            conditionalFormatDataNotAdd.Style.Fill.BackgroundColor.Color = Color.Red;

            #endregion

            for (int i = 0; i < data.resultFindGroups.Length; i++)
            {
                worksheet.Cells[startRow, startCol + 3 + i].Value = data.resultFindGroups[i].findStr;
                worksheet.Cells[startRow + 1, startCol + 3 + i].Value = GetStrFindStatGroup(data.resultFindGroups[i].resultFind);
                worksheet.Cells[startRow + 2, startCol + 3 + i].Value = data.resultFindGroups[i].groupCn;
            }

            for (int i = 0; i < data.resultFindUsers.Length; i++)
            {
                worksheet.Cells[startRow + 3 + i, startCol].Value = data.resultFindUsers[i].findStr;
                worksheet.Cells[startRow + 3 + i, startCol + 1].Value = GetStrFindStatUser(data.resultFindUsers[i].resultFind);
                worksheet.Cells[startRow + 3 + i, startCol + 2].Value = data.resultFindUsers[i].login;
            }

            for (int i = 0; i < data.resultUserInGroups.GetLength(0); i++)
            {
                for (int j = 0; j < data.resultUserInGroups.GetLength(1); j++)
                {
                    worksheet.Cells[startRow + dataOffsetRow + j, startCol + dataOffsetColl + i].Value = resultUserInGroupStatusToString[data.resultUserInGroups[i, j]];
                }
            }


            worksheet.Cells.AutoFitColumns();

            string path = $"{Path.GetTempFileName()}.xlsx";

            excelPackage.SaveAs(new FileInfo(path));

            Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE", path);
        }
    }
}