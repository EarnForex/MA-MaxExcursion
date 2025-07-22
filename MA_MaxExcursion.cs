// -------------------------------------------------------------------------------
//   MA MaxExcursion calculates maximum excursion of the price from its moving averages between two crosses.
//   Includes optional statistics calculation and alerts.
//   
//   Version 1.00
//   Copyright 2025, EarnForex.com
//   https://www.earnforex.com/metatrader-indicators/MA-MaxExcursion/
// -------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using cAlgo.API;
using cAlgo.API.Indicators;

namespace cAlgo
{
    [Indicator(IsOverlay = true, TimeZone = TimeZones.UTC, AccessRights = AccessRights.None)]
    public class MAMaxExcursion : Indicator
    {
        public enum ENUM_DISTANCE_MODE
        {
            Relative,
            Absolute
        };

        public enum ENUM_CROSS_POINT
        {
            Close_Price,
            MA_Value
        };

        // Main parameters
        [Parameter("MA Period", DefaultValue = 20, Group = "Main Settings")]
        public int MAPeriod { get; set; }

        [Parameter("MA Type", DefaultValue = MovingAverageType.Simple, Group = "Main Settings")]
        public MovingAverageType MAType { get; set; }

        [Parameter("Applied Price", Group = "Main Settings")]
        public DataSeries AppliedPrice { get; set; }

        [Parameter("Distance Mode (Absolute/Relative)", DefaultValue = ENUM_DISTANCE_MODE.Absolute, Group = "Main Settings")]
        public ENUM_DISTANCE_MODE UseAbsoluteMode { get; set; }

        [Parameter("Cross Point (MA Value/Close Price)", DefaultValue = ENUM_CROSS_POINT.MA_Value, Group = "Main Settings")]
        public ENUM_CROSS_POINT UseMAValueAtCross { get; set; }

        [Parameter("Font Size", DefaultValue = 10, Group = "Main Settings")]
        public int FontSize { get; set; }

        [Parameter("Label Color Above", DefaultValue = "Lime", Group = "Main Settings")]
        public Color LabelColorAbove { get; set; }

        [Parameter("Label Color Below", DefaultValue = "Red", Group = "Main Settings")]
        public Color LabelColorBelow { get; set; }

        // Processing parameters
        [Parameter("Max Bars to Process (0 = All)", DefaultValue = 1000, Group = "Processing Settings")]
        public int MaxBars { get; set; }

        // Alert parameters
        [Parameter("Enable Alerts", DefaultValue = false, Group = "Alert Settings")]
        public bool EnableAlerts { get; set; }

        [Parameter("Use Email Alerts", DefaultValue = false, Group = "Alert Settings")]
        public bool UseEmailAlert { get; set; }

        [Parameter("Sending Email Address", DefaultValue = "sender@example.com", Group = "Alert Settings")]
        public string SendingEmailAddress { get; set; }

        [Parameter("Receiving Email Address", DefaultValue = "receiver@example.com", Group = "Alert Settings")]
        public string ReceivingEmailAddress { get; set; }

        [Parameter("Use Sound Alert", DefaultValue = false, Group = "Alert Settings")]
        public bool UseSoundAlert { get; set; }

        [Parameter("Sound Type", DefaultValue = SoundType.Announcement, Group = "Alert Settings")]
        public SoundType SoundType { get; set; }

        // Statistics parameters
        [Parameter("Show Excursion Statistics", DefaultValue = true, Group = "Statistics Settings")]
        public bool ShowStatistics { get; set; }

        [Parameter("Statistics X Offset", DefaultValue = 10, Group = "Statistics Settings")]
        public int StatsXOffset { get; set; }

        [Parameter("Statistics Y Offset", DefaultValue = 50, Group = "Statistics Settings")]
        public int StatsYOffset { get; set; }

        [Parameter("Number of Recent Excursions (0 = All)", DefaultValue = 20, Group = "Statistics Settings")]
        public int StatsCount { get; set; }

        [Parameter("Statistics Text Color", DefaultValue = "White", Group = "Statistics Settings")]
        public Color StatsColor { get; set; }

        [Parameter("Statistics Font Size", DefaultValue = 9, Group = "Statistics Settings")]
        public int StatsFontSize { get; set; }

        // Indicator outputs
        [Output("Moving Average", LineColor = "Blue", LineStyle = LineStyle.Solid, Thickness = 2)]
        public IndicatorDataSeries MABuffer { get; set; }

        [Output("MA Excursion ZigZag", LineColor = "Yellow", LineStyle = LineStyle.Solid, Thickness = 2)]
        public IndicatorDataSeries ZigZagBuffer { get; set; }

        // Internal buffers (not plotted)
        public IndicatorDataSeries UpExcursionBuffer { get; set; }
        public IndicatorDataSeries DownExcursionBuffer { get; set; }

        // Private fields
        private MovingAverage _ma;
        private DateTime _lastAlertedCross = DateTime.MinValue;
        private DateTime _lastAddedExcursionTime = DateTime.MinValue;
        
        // Excursion tracking
        private List<double> _allExcursions = new List<double>();
        private List<double> _upExcursions = new List<double>();
        private List<double> _downExcursions = new List<double>();

        // UI Controls
        private TextBlock[] _statsLabels;

        // Cross tracking
        private struct ExcursionInfo
        {
            public DateTime CrossTime;
            public double ExcursionSize;
            public bool WasUpExcursion;
            public DateTime ExcursionTime;
        }

        protected override void Initialize()
        {
            // Initialize the moving average.
            _ma = Indicators.MovingAverage(AppliedPrice, MAPeriod, MAType);

            // Initialize internal buffers.
            UpExcursionBuffer = CreateDataSeries();
            DownExcursionBuffer = CreateDataSeries();

            // Create UI elements.
            if (ShowStatistics)
            {
                CreateStatisticsLabels();
            }
        }

        public override void Calculate(int index)
        {
            if (index < MAPeriod) return;

            // Calculate MA.
            MABuffer[index] = _ma.Result[index];

            // On last calculation run, perform full zigzag calculation
            if (IsLastBar && index == Bars.Count - 1)
            {
                CalculateZigZag();
            }
        }

        private void CalculateZigZag()
        {
            // Clear old chart objects.
            var objectsToRemove = Chart.Objects.Where(obj => obj.Name.StartsWith("MA_Excursion_")).ToList();
            foreach (var obj in objectsToRemove)
            {
                Chart.RemoveObject(obj.Name);
            }

            // Initialize buffers.
            int barsToProcess = MaxBars > 0 && MaxBars < Bars.Count ? MaxBars : Bars.Count;
            
            for (int i = 0; i < Bars.Count; i++)
            {
                ZigZagBuffer[i] = double.NaN;
                UpExcursionBuffer[i] = double.NaN;
                DownExcursionBuffer[i] = double.NaN;
            }

            // Reset excursion arrays
            _allExcursions.Clear();
            _upExcursions.Clear();
            _downExcursions.Clear();

            // Find crosses and excursions
            int lastCrossIndex = -1;
            bool lastWasCrossUp = false;

            // Process from oldest to newest bar (but limit to barsToProcess)
            int startIndex = Math.Max(MAPeriod + 1, Bars.Count - barsToProcess);
            
            for (int i = startIndex; i < Bars.Count - 1; i++)
            {
                double priceCurrentBar = GetPrice(i);
                double pricePrevBar = GetPrice(i - 1);
                double maCurrentBar = MABuffer[i];
                double maPrevBar = MABuffer[i - 1];

                bool isCross = false;
                bool isCrossUp = false;

                // Check for cross up
                if (pricePrevBar <= maPrevBar && priceCurrentBar > maCurrentBar)
                {
                    isCross = true;
                    isCrossUp = true;
                }
                // Check for cross down
                else if (pricePrevBar >= maPrevBar && priceCurrentBar < maCurrentBar)
                {
                    isCross = true;
                    isCrossUp = false;
                }

                if (isCross)
                {
                    // Mark cross point
                    ZigZagBuffer[i] = (UseMAValueAtCross == ENUM_CROSS_POINT.MA_Value) ? maCurrentBar : priceCurrentBar;

                    // If we have a previous cross, find excursion
                    if (lastCrossIndex >= 0)
                    {
                        var excursion = FindAndMarkExcursion(lastCrossIndex, i, lastWasCrossUp);
                        
                        if (excursion.ExcursionSize > 0)
                        {
                            _allExcursions.Add(excursion.ExcursionSize);
                            if (lastWasCrossUp)
                                _upExcursions.Add(excursion.ExcursionSize);
                            else
                                _downExcursions.Add(excursion.ExcursionSize);
                        }
                    }

                    lastCrossIndex = i;
                    lastWasCrossUp = isCrossUp;
                }
            }

            // Handle excursion from last cross to current bar
            if (lastCrossIndex >= 0 && lastCrossIndex < Bars.Count - 1)
            {
                var excursion = FindAndMarkExcursion(lastCrossIndex, Bars.Count - 1, lastWasCrossUp);
                
                if (excursion.ExcursionSize > 0 && excursion.ExcursionTime > DateTime.MinValue)
                {
                    _allExcursions.Add(excursion.ExcursionSize);
                    if (lastWasCrossUp)
                        _upExcursions.Add(excursion.ExcursionSize);
                    else
                        _downExcursions.Add(excursion.ExcursionSize);
                    
                    // Check for alerts on the most recent completed excursion
                    if (EnableAlerts && lastCrossIndex == Bars.Count - 2)
                    {
                        SendExcursionAlert(excursion);
                    }
                }
            }

            // Draw zigzag lines
            DrawZigZagLines();

            // Update statistics
            if (ShowStatistics)
            {
                UpdateStatistics();
            }
        }

        private void DrawZigZagLines()
        {
            List<int> zigzagPoints = new List<int>();
            
            // Collect all zigzag points
            for (int i = 0; i < Bars.Count; i++)
            {
                if (!double.IsNaN(ZigZagBuffer[i]))
                {
                    zigzagPoints.Add(i);
                }
            }
        }

        private ExcursionInfo FindAndMarkExcursion(int startIdx, int endIdx, bool fromCrossUp)
        {
            var info = new ExcursionInfo
            {
                CrossTime = Bars.OpenTimes[startIdx],
                ExcursionSize = 0,
                WasUpExcursion = fromCrossUp,
                ExcursionTime = DateTime.MinValue
            };

            double maxExcursion = 0;
            int maxExcursionIdx = -1;
            double maxPrice = 0;

            double crossValue = ZigZagBuffer[startIdx];

            if (UseAbsoluteMode == ENUM_DISTANCE_MODE.Absolute)
            {
                if (fromCrossUp)
                {
                    // Find highest point
                    double highestPrice = 0;
                    int highestIdx = -1;
                    double referenceValue = (UseMAValueAtCross == ENUM_CROSS_POINT.MA_Value) ? double.MaxValue : crossValue;

                    for (int k = startIdx; k <= endIdx; k++)
                    {
                        double highK = Bars.HighPrices[k];
                        if (highK > highestPrice)
                        {
                            highestPrice = highK;
                            highestIdx = k;
                        }
                        if (UseMAValueAtCross == ENUM_CROSS_POINT.MA_Value && MABuffer[k] < referenceValue)
                        {
                            referenceValue = MABuffer[k];
                        }
                    }

                    if (highestIdx >= 0)
                    {
                        maxExcursion = (UseMAValueAtCross == ENUM_CROSS_POINT.MA_Value) ? highestPrice - referenceValue : highestPrice - crossValue;
                        maxExcursionIdx = highestIdx;
                        maxPrice = highestPrice;
                    }
                }
                else
                {
                    // Find lowest point
                    double lowestPrice = double.MaxValue;
                    int lowestIdx = -1;
                    double referenceValue = (UseMAValueAtCross == ENUM_CROSS_POINT.MA_Value) ? 0 : crossValue;

                    for (int k = startIdx; k <= endIdx; k++)
                    {
                        double lowK = Bars.LowPrices[k];
                        if (lowK < lowestPrice)
                        {
                            lowestPrice = lowK;
                            lowestIdx = k;
                        }
                        if (UseMAValueAtCross == ENUM_CROSS_POINT.MA_Value && MABuffer[k] > referenceValue)
                        {
                            referenceValue = MABuffer[k];
                        }
                    }

                    if (lowestIdx >= 0)
                    {
                        maxExcursion = (UseMAValueAtCross == ENUM_CROSS_POINT.MA_Value) ? referenceValue - lowestPrice : crossValue - lowestPrice;
                        maxExcursionIdx = lowestIdx;
                        maxPrice = lowestPrice;
                    }
                }
            }
            else
            {
                // Relative mode
                for (int k = startIdx; k <= endIdx; k++)
                {
                    double distance = 0;
                    double priceExtreme = 0;

                    if (fromCrossUp)
                    {
                        double highK = Bars.HighPrices[k];
                        distance = (UseMAValueAtCross == ENUM_CROSS_POINT.MA_Value) ? highK - MABuffer[k] : highK - crossValue;
                        priceExtreme = highK;
                    }
                    else
                    {
                        double lowK = Bars.LowPrices[k];
                        distance = (UseMAValueAtCross == ENUM_CROSS_POINT.MA_Value) ? MABuffer[k] - lowK : crossValue - lowK;
                        priceExtreme = lowK;
                    }

                    if (distance > maxExcursion)
                    {
                        maxExcursion = distance;
                        maxExcursionIdx = k;
                        maxPrice = priceExtreme;
                    }
                }
            }

            // Mark the excursion point
            if (maxExcursionIdx >= 0 && maxExcursion > 0)
            {
                // Add to ZigZag buffer if not the start cross point
                if (maxExcursionIdx != startIdx)
                {
                    ZigZagBuffer[maxExcursionIdx] = maxPrice;
                }

                // Add text label
                string labelSuffix = fromCrossUp ? "_u" : "_d";
                string labelName = $"MA_Excursion_{Bars.OpenTimes[maxExcursionIdx].Ticks}{labelSuffix}";
                string distanceText = (maxExcursion / Symbol.PipSize).ToString("F1");

                Chart.DrawText(labelName, distanceText, maxExcursionIdx, maxPrice, 
                    fromCrossUp ? LabelColorAbove : LabelColorBelow);

                // Set buffer values
                if (fromCrossUp)
                {
                    UpExcursionBuffer[maxExcursionIdx] = maxExcursion / Symbol.PipSize;
                }
                else
                {
                    DownExcursionBuffer[maxExcursionIdx] = maxExcursion / Symbol.PipSize;
                }

                info.ExcursionSize = maxExcursion / Symbol.PipSize;
                info.ExcursionTime = Bars.OpenTimes[maxExcursionIdx];
            }

            return info;
        }

        private void CreateStatisticsLabels()
        {
            _statsLabels = new TextBlock[11]; // Header, count, total (3), up (3), down (3)
            
            for (int i = 0; i < _statsLabels.Length; i++)
            {
                _statsLabels[i] = new TextBlock
                {
                    ForegroundColor = StatsColor,
                    FontSize = StatsFontSize,
                    HorizontalAlignment = HorizontalAlignment.Left,
                    VerticalAlignment = VerticalAlignment.Top,
                    Margin = new Thickness(StatsXOffset, StatsYOffset + (i * 15), 0, 0)
                };
                Chart.AddControl(_statsLabels[i]);
            }

            // Set header font to bold
            _statsLabels[0].FontWeight = FontWeight.Bold;
            _statsLabels[0].FontSize = StatsFontSize + 1;
        }

        private void UpdateStatistics()
        {
            if (!ShowStatistics || _statsLabels == null) return;

            // Calculate statistics
            int countAll = StatsCount == 0 || StatsCount >= _allExcursions.Count ? _allExcursions.Count : StatsCount;
            
            double avgAll = 0, medianAll = 0;
            double avgUp = 0, medianUp = 0;
            double avgDown = 0, medianDown = 0;

            if (countAll > 0 && _allExcursions.Count > 0)
            {
                var recentAll = GetLastElements(_allExcursions, countAll);
                avgAll = recentAll.Average();
                medianAll = CalculateMedian(recentAll);
            }

            int countUp = Math.Min(countAll, _upExcursions.Count);
            if (countUp > 0)
            {
                var recentUp = GetLastElements(_upExcursions, countUp);
                avgUp = recentUp.Average();
                medianUp = CalculateMedian(recentUp);
            }

            int countDown = Math.Min(countAll, _downExcursions.Count);
            if (countDown > 0)
            {
                var recentDown = GetLastElements(_downExcursions, countDown);
                avgDown = recentDown.Average();
                medianDown = CalculateMedian(recentDown);
            }

            // Update labels
            _statsLabels[0].Text = "MA EXCURSION STATISTICS";
            _statsLabels[1].Text = $"({(StatsCount == 0 ? $"All {_allExcursions.Count}" : $"{countAll} Recent")} Excursions)";
            _statsLabels[1].FontSize = StatsFontSize - 1;
            
            _statsLabels[2].Text = "TOTAL:";
            _statsLabels[2].FontWeight = FontWeight.Bold;
            _statsLabels[3].Text = $"  Avg: {avgAll:F1} points";
            _statsLabels[4].Text = $"  Med: {medianAll:F1} points";
            
            _statsLabels[5].Text = "UP:";
            _statsLabels[5].FontWeight = FontWeight.Bold;
            _statsLabels[5].ForegroundColor = LabelColorAbove;
            _statsLabels[6].Text = $"  Avg: {avgUp:F1} points";
            _statsLabels[6].ForegroundColor = LabelColorAbove;
            _statsLabels[7].Text = $"  Med: {medianUp:F1} points";
            _statsLabels[7].ForegroundColor = LabelColorAbove;
            
            _statsLabels[8].Text = "DOWN:";
            _statsLabels[8].FontWeight = FontWeight.Bold;
            _statsLabels[8].ForegroundColor = LabelColorBelow;
            _statsLabels[9].Text = $"  Avg: {avgDown:F1} points";
            _statsLabels[9].ForegroundColor = LabelColorBelow;
            _statsLabels[10].Text = $"  Med: {medianDown:F1} points";
            _statsLabels[10].ForegroundColor = LabelColorBelow;
        }

        private List<double> GetLastElements(List<double> list, int count)
        {
            if (list.Count <= count)
                return new List<double>(list);
            
            var result = new List<double>();
            for (int i = list.Count - count; i < list.Count; i++)
            {
                result.Add(list[i]);
            }
            return result;
        }

        private double CalculateMedian(List<double> values)
        {
            if (values.Count == 0) return 0;
            
            var sorted = values.OrderBy(v => v).ToList();
            int mid = sorted.Count / 2;
            
            if (sorted.Count % 2 == 0)
            {
                return (sorted[mid - 1] + sorted[mid]) / 2.0;
            }
            else
            {
                return sorted[mid];
            }
        }

        private double GetPrice(int index)
        {
            return AppliedPrice[index];
        }

        private void SendExcursionAlert(ExcursionInfo excursion)
        {
            // Don't alert for the same cross twice.
            if (excursion.CrossTime == _lastAlertedCross) return;
            
            _lastAlertedCross = excursion.CrossTime;
            
            var direction = excursion.WasUpExcursion ? "UP" : "DOWN";
            var message = $"MA Cross Excursion: {Symbol.Name} {TimeFrame} | {direction} excursion: {excursion.ExcursionSize:F1} points";

            if (UseEmailAlert)
            {
                Notifications.SendEmail(SendingEmailAddress, ReceivingEmailAddress, "MA Cross Excursion Alert", message);
            }

            if (UseSoundAlert)
            {
                Notifications.PlaySound(SoundType);
            }
        }
    }
}