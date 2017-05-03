using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace Challenge
{
    public partial class MainWindow : Form
    {
        string jsonPath = null;
        public MainWindow()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            List<Challenge> chalFromJson = JsonConvert.DeserializeObject<List<Challenge>>(File.ReadAllText(@jsonPath));
            Stream myStream = null;
            string startupPath = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                button1.Enabled = false;
                button2.Enabled = false;
                label1.Text = "The task is in process...";
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            Daily daily = null;                                     //object with numbers of daily challenges
                            int prevDay = 0;
                            int day = 0;
                            int nextday = 0;                                        //variable for separeting daily files
                            string reqMess = null;
                            string reqMessLow = null;

                            bool addNewFlag = false;                                //var for adding new challenge in json
                            int newId = chalFromJson.Count - 1;                     // Id for new challenge
                            Excel.Application xlApp = new Excel.Application();
                            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@openFileDialog1.FileName);
                            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                            Excel.Range xlRange = xlWorksheet.UsedRange;
                            int rowCount = xlRange.Rows.Count;
                            for (int i = 2; i <= rowCount; i++)
                            {
                                Challenge chalFromXlsx = new Challenge();
                                reqMess = xlRange.Cells[i, 4].Value.ToString();
                                day = (int)xlRange.Cells[i, 1].Value;
                                if (day > prevDay)                                  //creating new daily folder and object
                                {
                                    startupPath = System.IO.Directory.GetCurrentDirectory() + "\\Challenges\\";
                                    System.IO.Directory.CreateDirectory(System.IO.Path.Combine(startupPath, day.ToString()));
                                    daily = new Daily();
                                    daily.period = 86400000;
                                    if (daily.availableIds == null)
                                    {
                                        daily.availableIds = new List<int>();
                                    }
                                    prevDay++;
                                }
                                #region Make objects from xls
                                //switch (reqMess)
                                //{
                                //    case "Play X spins.":
                                //        chalFromXlsx.Name = "Can't Stop x%X";
                                //        chalFromXlsx.RequirementId = "play_x_spins";
                                //        chalFromXlsx.RequirementMsg = "Play %X spins.";
                                //        if (xlRange.Cells[i, 7].Value.ToString() == "Credits")
                                //        {
                                //            chalFromXlsx.RewardType = "coins";
                                //        }
                                //        chalFromXlsx.RewardAmount = (long)xlRange.Cells[i, 8].Value;
                                //        chalFromXlsx.CounterName = "spins";
                                //        chalFromXlsx.CounterGoal = (long)xlRange.Cells[i, 5].Value;
                                //        chalFromXlsx.SlotsCondition = "any";
                                //        break;
                                //    case "Complete all other daily challenges.":
                                //        chalFromXlsx.Name = "Perfect Score";
                                //        chalFromXlsx.RequirementId = "complete_all_other_daily_challenges";
                                //        chalFromXlsx.RequirementMsg = "Complete all other daily challenges.";
                                //        if (xlRange.Cells[i, 7].Value.ToString() == "Credits")
                                //        {
                                //            chalFromXlsx.RewardType = "coins";
                                //        }
                                //        chalFromXlsx.RewardAmount = (long)xlRange.Cells[i, 8].Value;
                                //        chalFromXlsx.CounterName = "completedChallenges";
                                //        chalFromXlsx.SlotsCondition = "any";
                                //        break;
                                //    case "Bet a total of Y credits.":
                                //        chalFromXlsx.Name = "Taking Charge";
                                //        chalFromXlsx.RequirementId = "bet_a_total_of_x_coins";
                                //        chalFromXlsx.RequirementMsg = "Bet a total of %X coins.";
                                //        if (xlRange.Cells[i, 7].Value.ToString() == "Credits")
                                //        {
                                //            chalFromXlsx.RewardType = "coins";
                                //        }
                                //        chalFromXlsx.RewardAmount = (long)xlRange.Cells[i, 8].Value;
                                //        chalFromXlsx.CounterName = "bet";
                                //        chalFromXlsx.CounterGoal = (long)xlRange.Cells[i, 6].Value;
                                //        chalFromXlsx.SlotsCondition = "any";
                                //        break;
                                //    case "Win credits in X different slots.":
                                //        chalFromXlsx.Name = "Variety Pack";
                                //        chalFromXlsx.RequirementId = "win_credits_in_x_different_slots";
                                //        chalFromXlsx.RequirementMsg = "Win coins in %X different slots.";
                                //        if (xlRange.Cells[i, 7].Value.ToString() == "Credits")
                                //        {
                                //            chalFromXlsx.RewardType = "coins";
                                //        }
                                //        chalFromXlsx.RewardAmount = (long)xlRange.Cells[i, 8].Value;
                                //        chalFromXlsx.CounterName = "spins";
                                //        chalFromXlsx.CounterGoal = (long)xlRange.Cells[i, 5].Value;
                                //        chalFromXlsx.ConditionName = "winsAmount";
                                //        chalFromXlsx.ConditionValue = 1;
                                //        chalFromXlsx.SlotsCondition = "unique";
                                //        break;
                                //    case "Win a total of Y credits.":
                                //        chalFromXlsx.Name = "Winner's Reward";
                                //        chalFromXlsx.RequirementId = "win_a_total_of_x_credits";
                                //        chalFromXlsx.RequirementMsg = "Win a total of %X coins.";
                                //        if (xlRange.Cells[i, 7].Value.ToString() == "Credits")
                                //        {
                                //            chalFromXlsx.RewardType = "coins";
                                //        }
                                //        chalFromXlsx.RewardAmount = (long)xlRange.Cells[i, 8].Value;
                                //        chalFromXlsx.CounterName = "winsAmount";
                                //        chalFromXlsx.CounterGoal = (long)xlRange.Cells[i, 6].Value;
                                //        chalFromXlsx.SlotsCondition = "any";
                                //        break;
                                //    case "Get X wins in a row.":
                                //        chalFromXlsx.Name = "Winning Streak x%X";
                                //        chalFromXlsx.RequirementId = "get_x_wins_in_a_row_on_any_slot";
                                //        chalFromXlsx.RequirementMsg = "Get %X wins in a row on any slot.";
                                //        if (xlRange.Cells[i, 7].Value.ToString() == "Credits")
                                //        {
                                //            chalFromXlsx.RewardType = "coins";
                                //        }
                                //        chalFromXlsx.RewardAmount = (long)xlRange.Cells[i, 8].Value;
                                //        chalFromXlsx.CounterName = "winsInARow";
                                //        chalFromXlsx.CounterGoal = (long)xlRange.Cells[i, 5].Value;
                                //        chalFromXlsx.SlotsCondition = "any";
                                //        break;
                                //    default:
                                //        break;
                                //}
                                #endregion
                                #region Make objects from xls V2
                                /////////////////////////////////////////////////////////////////
                                reqMessLow = reqMess.ToLower().Replace(" ", "_").TrimEnd('.');
                                if (reqMessLow == "bet_a_total_of_y_credits")
                                {
                                    reqMessLow = "bet_a_total_of_x_coins";
                                }
                                if (reqMessLow == "win_a_total_of_y_credits")
                                {
                                    reqMessLow = "win_a_total_of_x_credits";
                                }
                                if (reqMessLow == "get_x_wins_in_a_row")
                                {
                                    reqMessLow = "get_x_wins_in_a_row_on_any_slot";
                                }
                                if (reqMessLow == "bet_y_or_more_on_x_different_slots_each")
                                {
                                    reqMessLow = "bet_y_on_x_different_slots_each";
                                }
                                if (reqMessLow == "win_y_or_more_in_a_spin_x_times")
                                {
                                    reqMessLow = "win_y_in_x_spins";
                                }
                                foreach (var item in chalFromJson)              //comparing object from xlsx with object from json
                                {
                                    if (item.requirementId == reqMessLow)
                                    {
                                        chalFromXlsx.name = item.name;
                                        //
                                        chalFromXlsx.requirementId = item.requirementId;
                                        //
                                        chalFromXlsx.requirementMsg = item.requirementMsg;
                                        //
                                        if (xlRange.Cells[i, 7].Value.ToString() == "Credits")
                                        {
                                            chalFromXlsx.rewardType = "coins";
                                        }
                                        if (xlRange.Cells[i, 7].Value.ToString() == "VIP")
                                        {
                                            chalFromXlsx.rewardType = "vp";
                                        }
                                        //
                                        chalFromXlsx.rewardAmount = (long)xlRange.Cells[i, 8].Value;
                                        //
                                        chalFromXlsx.counterName = item.counterName;
                                        //
                                        if (item.counterGoal != 0)
                                        {
                                            if ((xlRange.Cells[i, 6].Value == null) || ((long)xlRange.Cells[i, 6].Value == 0))
                                            {
                                                chalFromXlsx.counterGoal = (long)xlRange.Cells[i, 5].Value;
                                            }
                                            else if ((xlRange.Cells[i, 5].Value == null) || ((long)xlRange.Cells[i, 5].Value == 0))
                                            {
                                                chalFromXlsx.counterGoal = (long)xlRange.Cells[i, 6].Value;
                                            }
                                            else
                                            {
                                                chalFromXlsx.counterGoal = (long)xlRange.Cells[i, 5].Value;
                                            }
                                        }
                                        //
                                        if (item.conditionName != null)
                                        {
                                            chalFromXlsx.conditionName = item.conditionName;
                                        }
                                        //
                                        if (item.conditionValue == 0)
                                        {
                                            chalFromXlsx.conditionValue = 0;
                                        }
                                        else if (item.conditionValue == 1)
                                        {
                                            chalFromXlsx.conditionValue = 1;
                                        }
                                        else
                                        {
                                            chalFromXlsx.conditionValue = (long)xlRange.Cells[i, 6].Value;
                                        }
                                        //
                                        chalFromXlsx.slotsCondition = item.slotsCondition;
                                        break;
                                    }

                                }
                                /////////////////////////////////////////////////////////////
                                #endregion //v2
                                foreach (var item in chalFromJson)              //comparing object from xlsx with object from json
                                {
                                    if (item.Equals(chalFromXlsx) == true)
                                    {
                                        daily.availableIds.Add(item.id);        //adding to daily obj
                                        break;
                                    }
                                    else
                                    {
                                        if (chalFromJson.IndexOf(item) == chalFromJson.Count - 1)
                                        {
                                            addNewFlag = true;                  //flag for adding to json obj new object
                                        }
                                    }
                                }
                                if (addNewFlag)                                 //adding to json obj new object
                                {
                                    chalFromXlsx.id = ++newId;
                                    chalFromJson.Add(chalFromXlsx);
                                    daily.availableIds.Add(chalFromXlsx.id);
                                    addNewFlag = false;
                                }
                                nextday = (xlRange.Cells[i + 1, 1].Value == null) ? 15 : (int)xlRange.Cells[i + 1, 1].Value;
                                if (day < nextday)
                                {
                                    daily.availableIds.Sort();
                                    File.WriteAllText(@startupPath + day.ToString() + "\\availableChallengesConfigs.json", JsonConvert.SerializeObject(daily, Formatting.Indented));
                                }
                            }
                            #region clear memory
                            //cleanup
                            GC.Collect();
                            GC.WaitForPendingFinalizers();

                            //release com objects to fully kill excel process from running in the background
                            Marshal.ReleaseComObject(xlRange);
                            Marshal.ReleaseComObject(xlWorksheet);

                            //close and release
                            xlWorkbook.Close();
                            Marshal.ReleaseComObject(xlWorkbook);

                            //quit and release
                            xlApp.Quit();
                            Marshal.ReleaseComObject(xlApp);
                            #endregion
                            File.WriteAllText(@startupPath + "\\challengesConfigs.json", JsonConvert.SerializeObject(chalFromJson, Formatting.Indented));
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Original error: " + ex.Message);
                }
            }
            MessageBox.Show("Done! Created config files in " + @startupPath, "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Application.Exit();
        }
        private void button2_Click(object sender, EventArgs e)
        {

            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
            openFileDialog2.Filter = "json files (*.json)|*.json|All files (*.*)|*.*";
            openFileDialog2.FilterIndex = 1;
            openFileDialog2.RestoreDirectory = true;
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                jsonPath = @openFileDialog2.FileName;
                button1.Enabled = true;
            }
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
            button1.Enabled = false;

        }
    }
}
