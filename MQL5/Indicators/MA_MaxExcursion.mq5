//+------------------------------------------------------------------+
//|                                              MA MaxExcursion.mq5 |
//|                                      Copyright © 2025, EarnForex |
//|                                        https://www.earnforex.com |
//+------------------------------------------------------------------+
#property copyright "www.EarnForex.com, 2025"
#property link      "https://www.earnforex.com/metatrader-indicators/MA-MaxExcursion/"
#property version   "1.00"
#property icon      "\\Files\\EF-Icon-64x64px.ico"

#property description "MA MaxExcursion calculates maximum excursion of the price from its moving averages between two crosses."
#property description "Includes optional statistics calculation and alerts."

#property indicator_chart_window
#property indicator_buffers 4
#property indicator_plots   4

// MA line
#property indicator_label1  "MA"
#property indicator_type1   DRAW_LINE
#property indicator_color1  clrBlue
#property indicator_style1  STYLE_SOLID
#property indicator_width1  2

// Section line for ZigZag pattern
#property indicator_label2  "MA Excursion ZigZag"
#property indicator_type2   DRAW_SECTION
#property indicator_color2  clrYellow
#property indicator_style2  STYLE_SOLID
#property indicator_width2  2

// Up excursion values (hidden plot)
#property indicator_label3  "Up Excursion"
#property indicator_type3   DRAW_NONE

// Down excursion values (hidden plot)
#property indicator_label4  "Down Excursion"
#property indicator_type4   DRAW_NONE

// MA line
#property indicator_label1  "MA"
#property indicator_type1   DRAW_LINE
#property indicator_color1  clrBlue
#property indicator_style1  STYLE_SOLID
#property indicator_width1  2

// Section line for ZigZag pattern
#property indicator_label2  "MA Excursion ZigZag"
#property indicator_type2   DRAW_SECTION
#property indicator_color2  clrYellow
#property indicator_style2  STYLE_SOLID
#property indicator_width2  2

enum ENUM_DISTANCE_MODE
{
    Relative,
    Absolute
};

enum ENUM_CROSS_POINT
{
    Close_Price, // Close Price
    MA_Value // MA Value
};

// Input parameters
input string             MainSection = "=== Main Settings ==="; // Main Settings
input int                MA_Period = 20;            // MA Period
input ENUM_MA_METHOD     MA_Method = MODE_SMA;      // MA Method
input ENUM_APPLIED_PRICE MA_Price = PRICE_CLOSE;    // Applied Price
input ENUM_DISTANCE_MODE UseAbsoluteMode = Absolute;// Distance Mode
input ENUM_CROSS_POINT   UseMAValueAtCross = MA_Value; // Cross Point
input int                FontSize = 10;             // Font Size for Labels
input color              LabelColorAbove = clrLime; // Label Color Above MA
input color              LabelColorBelow = clrRed;  // Label Color Below MA
// Processing parameters
input string             ProcessSection = "=== Processing Settings ==="; // Processing Settings
input int                MaxBars = 1000;            // Max Bars to Process (0 = All)
// Alert parameters
input string             AlertSection = "=== Alert Settings ==="; // Alert Settings
input bool               EnableAlerts = false;      // Enable Alerts
input bool               UsePopupAlert = true;      // Use Popup Alerts
input bool               UseEmailAlert = false;     // Use Email Alerts
input bool               UsePushAlert = false;      // Use Push Notifications
input bool               UseSoundAlert = false;     // Use Sound Alert
input string             SoundFile = "alert.wav";   // Sound File Name
// Statistics parameters
input string             StatsSection = "=== Statistics Settings ==="; // Statistics Settings
input bool               ShowStatistics = true;     // Show Excursion Statistics
input ENUM_BASE_CORNER   StatsCorner = CORNER_LEFT_UPPER; // Statistics Corner
input int                StatsXOffset = 0;          // Statistics X Offset (+/-)
input int                StatsCount = 20;           // Number of Recent Excursions (0 = All)
input color              StatsColor = clrWhite;     // Statistics Text Color
input int                StatsFontSize = 9;         // Statistics Font Size
// Button parameters
input string             ButtonSection = "=== Button Settings ==="; // Button Settings
input ENUM_BASE_CORNER   ButtonCorner = CORNER_RIGHT_UPPER; // Button Corner
input int                ButtonXOffset = 100;       // Button X Offset
input int                ButtonYOffset = 10;        // Button Y Offset

// Indicator buffers
double MABuffer[];
double ZigZagBuffer[];
double UpExcursionBuffer[];
double DownExcursionBuffer[];

struct ExcursionInfo
{
    datetime cross_time;
    double excursion_size;
    bool was_up_excursion;
    datetime excursion_time;
};
ExcursionInfo lastExcursion;

// Global variables for alerts:
datetime lastAlertedCross = 0;

// Arrays to store excursion history for statistics:
double AllExcursions[];
double UpExcursions[];
double DownExcursions[];
int AllExcursionsCount = 0;
int UpExcursionsCount = 0;
int DownExcursionsCount = 0;
datetime lastAddedExcursionTime = 0;  // Track last added excursion to avoid duplicates.
int ma_handle;

// Button control variables:
bool IndicatorEnabled = true;
string ButtonName = "MA_Exc_Button"; // The name is different to avoid deleting it by prefix.

int OnInit()
{
    // Set MA buffer
    SetIndexBuffer(0, MABuffer, INDICATOR_DATA);
    PlotIndexSetInteger(0, PLOT_DRAW_TYPE, DRAW_LINE);
    PlotIndexSetInteger(0, PLOT_LINE_COLOR, clrBlue);
    PlotIndexSetInteger(0, PLOT_LINE_WIDTH, 2);
    PlotIndexSetString(0, PLOT_LABEL, "Moving Average");

    // Set ZigZag buffer using DRAW_SECTION with 0.0 as empty value
    SetIndexBuffer(1, ZigZagBuffer, INDICATOR_DATA);
    PlotIndexSetInteger(1, PLOT_DRAW_TYPE, DRAW_SECTION);
    PlotIndexSetString(1, PLOT_LABEL, "MA Excursion ZigZag");
    PlotIndexSetDouble(1, PLOT_EMPTY_VALUE, 0.0);

    // Set Up Excursion buffer (hidden plot for iCustom access)
    SetIndexBuffer(2, UpExcursionBuffer, INDICATOR_DATA);
    PlotIndexSetInteger(2, PLOT_DRAW_TYPE, DRAW_NONE);
    PlotIndexSetString(2, PLOT_LABEL, "Up Excursion");
    PlotIndexSetDouble(2, PLOT_EMPTY_VALUE, EMPTY_VALUE);

    // Set Down Excursion buffer (hidden plot for iCustom access)
    SetIndexBuffer(3, DownExcursionBuffer, INDICATOR_DATA);
    PlotIndexSetInteger(3, PLOT_DRAW_TYPE, DRAW_NONE);
    PlotIndexSetString(3, PLOT_LABEL, "Down Excursion");
    PlotIndexSetDouble(3, PLOT_EMPTY_VALUE, EMPTY_VALUE);

    // Create MA handle:
    ma_handle = iMA(_Symbol, _Period, MA_Period, 0, MA_Method, MA_Price);
    if (ma_handle == INVALID_HANDLE)
    {
        Print("Error creating MA handle");
        return INIT_FAILED;
    }

    // Set arrays as series:
    ArraySetAsSeries(MABuffer, true);
    ArraySetAsSeries(ZigZagBuffer, true);
    ArraySetAsSeries(UpExcursionBuffer, true);
    ArraySetAsSeries(DownExcursionBuffer, true);

    // Create on/off button:
    CreateButton();

    return INIT_SUCCEEDED;
}

void OnDeinit(const int reason)
{
    // Delete all objects created by this indicator (except the button).
    ObjectsDeleteAll(0, "MA_Excursion_");
    // Delete button:
    ObjectDelete(0, ButtonName);
}

int OnCalculate(const int rates_total,
                const int prev_calculated,
                const datetime &time[],
                const double &open[],
                const double &high[],
                const double &low[],
                const double &close[],
                const long &tick_volume[],
                const long &volume[],
                const int &spread[])
{
    // Check for minimum bars:
    if (rates_total < MA_Period + 1) return 0;

    // If indicator is disabled, clear all buffers and return:
    if (!IndicatorEnabled)
    {
        // Clear all buffers:
        for (int i = 0; i < rates_total; i++)
        {
            MABuffer[i] = EMPTY_VALUE;
            ZigZagBuffer[i] = 0.0;
            UpExcursionBuffer[i] = EMPTY_VALUE;
            DownExcursionBuffer[i] = EMPTY_VALUE;
        }
        // Clear all text objects except the button.
        ObjectsDeleteAll(0, "MA_Excursion_");
        return 0;
    }

    // Determine how many bars to process:
    int bars_to_process = rates_total;
    if (MaxBars > 0 && MaxBars < rates_total)
    {
        bars_to_process = MaxBars;
    }

    // Copy MA values:
    int copied = CopyBuffer(ma_handle, 0, 0, bars_to_process, MABuffer);
    if (copied < bars_to_process)
    {
        Print("Failed to copy MA buffer");
        return 0;
    }

    if (prev_calculated == 0) // Initial calculation.
    {
        // Initialize buffers - set ZigZag to 0.0, others to EMPTY_VALUE.
        for (int i = 0; i < rates_total; i++)
        {
            MABuffer[i] = EMPTY_VALUE;
            ZigZagBuffer[i] = 0.0;
            UpExcursionBuffer[i] = EMPTY_VALUE;
            DownExcursionBuffer[i] = EMPTY_VALUE;
        }

        // Full recalculation on first run.
        CalculateZigZag(bars_to_process);
    }
    else
    {
        // Partial update for new bars/ticks.
        UpdateZigZag(bars_to_process, prev_calculated);

        // Update statistics display if enabled.
        if (ShowStatistics)
        {
            DisplayExcursionStatistics();
        }
    }

    return rates_total;
}

void OnChartEvent(const int id, const long &lparam, const double &dparam, const string &sparam)
{
    // Check if the on/off button was clicked.
    if (id == CHARTEVENT_OBJECT_CLICK && sparam == ButtonName)
    {
        // Toggle indicator state.
        IndicatorEnabled = !IndicatorEnabled;

        // If indicator is being disabled, immediately delete all objects.
        if (!IndicatorEnabled)
        {
            ObjectsDeleteAll(0, "MA_Excursion_");
        }

        ChartSetSymbolPeriod(0, NULL, 0); // Forces refresh.
        
        // Update button appearance
        UpdateButton();

        // Reset button state.
        ObjectSetInteger(0, ButtonName, OBJPROP_STATE, false);
        ChartRedraw();
    }
}

//+------------------------------------------------------------------+
//| Update ZigZag for new bars/ticks.                                |
//+------------------------------------------------------------------+
void UpdateZigZag(int bars_to_process, int prev_calculated)
{
    // Find the last completed cross in the ZigZag buffer.
    int last_cross_index = -1;
    bool last_was_cross_up = false;

    // Search for the most recent cross point (excluding current bar).
    for (int i = 1; i < bars_to_process - MA_Period - 1; i++)
    {
        if (ZigZagBuffer[i] != 0.0)
        {
            // Check if this is a cross point (not an excursion point).
            double price_current = iClose(_Symbol, _Period, i);
            double price_prev = iClose(_Symbol, _Period, i + 1);
            double ma_current = MABuffer[i];
            double ma_prev = MABuffer[i + 1];

            bool is_cross_up = (price_prev <= ma_prev && price_current > ma_current);
            bool is_cross_down = (price_prev >= ma_prev && price_current < ma_current);

            if (is_cross_up || is_cross_down)
            {
                last_cross_index = i;
                last_was_cross_up = is_cross_up;
                break;
            }
        }
    }

    // If we found a last cross, update only the excursion from that cross to current.
    if (last_cross_index >= 0)
    {
        // First, find if there's a previous cross before the last one.
        int prev_cross_index = -1;
        bool prev_was_cross_up = false;

        for (int i = last_cross_index + 1; i < bars_to_process - MA_Period - 1; i++)
        {
            if (ZigZagBuffer[i] != 0.0)
            {
                double price_current = iClose(_Symbol, _Period, i);
                double price_prev = iClose(_Symbol, _Period, i + 1);
                double ma_current = MABuffer[i];
                double ma_prev = MABuffer[i + 1];

                bool is_cross_up = (price_prev <= ma_prev && price_current > ma_current);
                bool is_cross_down = (price_prev >= ma_prev && price_current < ma_current);

                if (is_cross_up || is_cross_down)
                {
                    prev_cross_index = i;
                    prev_was_cross_up = is_cross_up;
                    break;
                }
            }
        }

        // Only clear ZigZag values that might change (from last cross onwards).
        // But preserve the cross point value to avoid blinking.
        double cross_value = ZigZagBuffer[last_cross_index];
        for (int i = last_cross_index; i >= 0; i--)
        {
            if (i != last_cross_index || ZigZagBuffer[i] != cross_value)
            {
                ZigZagBuffer[i] = 0.0;
            }
            // Clear excursion buffers.
            UpExcursionBuffer[i] = EMPTY_VALUE;
            DownExcursionBuffer[i] = EMPTY_VALUE;
        }

        // Remove text labels in the update range (including the cross bar).
        for (int i = last_cross_index; i >= 0; i--)
        {
            datetime bar_time = iTime(_Symbol, _Period, i);
            string label_base = "MA_Excursion_" + TimeToString(bar_time, TIME_DATE | TIME_MINUTES);

            // Remove both up and down labels if they exist.
            string label_up = label_base + "_u";
            string label_down = label_base + "_d";

            ObjectDelete(0, label_up);
            ObjectDelete(0, label_down);
        }

        // If there was a previous cross, recalculate the excursion that ENDS at last_cross.
        if (prev_cross_index >= 0)
        {
            ExcursionInfo excursion = FindAndMarkExcursion(prev_cross_index, last_cross_index, prev_was_cross_up);
        }

        // Recalculate excursion from last cross to current bar.
        FindAndMarkExcursion(last_cross_index, 0, last_was_cross_up);

        // Check if a new cross has formed on the previous bar (not current bar).
        for (int i = MathMin(2, last_cross_index - 1); i >= 1; i--)
        {
            double price_current = iClose(_Symbol, _Period, i);
            double price_prev = iClose(_Symbol, _Period, i + 1);
            double ma_current = MABuffer[i];
            double ma_prev = MABuffer[i + 1];

            bool is_cross_up = (price_prev <= ma_prev && price_current > ma_current);
            bool is_cross_down = (price_prev >= ma_prev && price_current < ma_current);

            if (is_cross_up || is_cross_down)
            {
                // New cross found, mark it.
                ZigZagBuffer[i] = (UseMAValueAtCross == MA_Value) ? ma_current : price_current;

                // Find excursion from previous cross to this new cross.
                ExcursionInfo excursion = FindAndMarkExcursion(last_cross_index, i, last_was_cross_up);

                // Add completed excursion to statistics.
                if (excursion.excursion_size > 0 && excursion.cross_time != lastAddedExcursionTime)
                {
                    ArrayResize(AllExcursions, AllExcursionsCount + 1);
                    AllExcursions[AllExcursionsCount] = excursion.excursion_size;
                    AllExcursionsCount++;

                    if (last_was_cross_up)
                    {
                        ArrayResize(UpExcursions, UpExcursionsCount + 1);
                        UpExcursions[UpExcursionsCount] = excursion.excursion_size;
                        UpExcursionsCount++;
                    }
                    else
                    {
                        ArrayResize(DownExcursions, DownExcursionsCount + 1);
                        DownExcursions[DownExcursionsCount] = excursion.excursion_size;
                        DownExcursionsCount++;
                    }

                    lastAddedExcursionTime = excursion.cross_time;
                }

                // Check if we should alert about the completed excursion.
                if (EnableAlerts && i == 1 && excursion.cross_time > 0 && excursion.cross_time != lastAlertedCross)
                {
                    SendExcursionAlert(excursion);
                    lastAlertedCross = excursion.cross_time;
                }

                // Update for next iteration.
                last_cross_index = i;
                last_was_cross_up = is_cross_up;
            }
        }
    }
    else
    {
        // No crosses found, do full recalculation.
        CalculateZigZag(bars_to_process);
    }
}

//+------------------------------------------------------------------+
//| Calculate complete ZigZag using DRAW_SECTION.                    |
//+------------------------------------------------------------------+
void CalculateZigZag(int bars_to_process)
{
    // Clear old text objects.
    ObjectsDeleteAll(0, "MA_Excursion_");

    // Initialize only the bars we'll process.
    for (int i = 0; i < bars_to_process; i++)
    {
        ZigZagBuffer[i] = 0.0;
        UpExcursionBuffer[i] = EMPTY_VALUE;
        DownExcursionBuffer[i] = EMPTY_VALUE;
    }

    // Set ZigZag buffer to 0.0 for bars before MaxBars.
    if (MaxBars > 0 && MaxBars < bars_to_process)
    {
        for (int i = MaxBars; i < bars_to_process; i++)
        {
            ZigZagBuffer[i] = 0.0;
        }
    }

    // Reset excursion arrays:
    ArrayResize(AllExcursions, 0);
    ArrayResize(UpExcursions, 0);
    ArrayResize(DownExcursions, 0);
    AllExcursionsCount = 0;
    UpExcursionsCount = 0;
    DownExcursionsCount = 0;

    // Variables to track crosses:
    int last_cross_index = -1;
    bool last_was_cross_up = false;

    // Find crosses and excursions.
    // Process from oldest to newest bar.
    for (int i = bars_to_process - MA_Period - 2; i >= 1; i--)
    {
        double price_current = iClose(_Symbol, _Period, i);
        double price_prev = iClose(_Symbol, _Period, i + 1);
        double ma_current = MABuffer[i];
        double ma_prev = MABuffer[i + 1];

        bool is_cross = false;
        bool is_cross_up = false;

        // Check for cross up (price crosses above MA).
        if (price_prev <= ma_prev && price_current > ma_current)
        {
            is_cross = true;
            is_cross_up = true;
        }
        // Check for cross down (price crosses below MA).
        else if (price_prev >= ma_prev && price_current < ma_current)
        {
            is_cross = true;
            is_cross_up = false;
        }

        // If we found a cross:
        if (is_cross)
        {
            // Mark cross point at MA value or Close price based on setting.
            ZigZagBuffer[i] = (UseMAValueAtCross == MA_Value) ? ma_current : price_current;

            // If we have a previous cross (which is actually older since we're going backwards), find excursion from that older cross to this newer cross.
            if (last_cross_index >= 0)
            {
                ExcursionInfo excursion = FindAndMarkExcursion(last_cross_index, i, last_was_cross_up);

                // Store excursion data for statistics.
                if (excursion.excursion_size > 0)
                {
                    ArrayResize(AllExcursions, AllExcursionsCount + 1);
                    AllExcursions[AllExcursionsCount] = excursion.excursion_size;
                    AllExcursionsCount++;

                    if (last_was_cross_up)
                    {
                        ArrayResize(UpExcursions, UpExcursionsCount + 1);
                        UpExcursions[UpExcursionsCount] = excursion.excursion_size;
                        UpExcursionsCount++;
                    }
                    else
                    {
                        ArrayResize(DownExcursions, DownExcursionsCount + 1);
                        DownExcursions[DownExcursionsCount] = excursion.excursion_size;
                        DownExcursionsCount++;
                    }
                }
            }

            // Update last cross info.
            last_cross_index = i;
            last_was_cross_up = is_cross_up;
        }
    }

    // Handle excursion from last cross to current bar.
    if (last_cross_index >= 0 && last_cross_index > 0)
    {
        ExcursionInfo excursion = FindAndMarkExcursion(last_cross_index, 0, last_was_cross_up);

        // Store excursion data for statistics (if completed).
        if (excursion.excursion_size > 0 && excursion.excursion_time > 0)
        {
            ArrayResize(AllExcursions, AllExcursionsCount + 1);
            AllExcursions[AllExcursionsCount] = excursion.excursion_size;
            AllExcursionsCount++;

            if (last_was_cross_up)
            {
                ArrayResize(UpExcursions, UpExcursionsCount + 1);
                UpExcursions[UpExcursionsCount] = excursion.excursion_size;
                UpExcursionsCount++;
            }
            else
            {
                ArrayResize(DownExcursions, DownExcursionsCount + 1);
                DownExcursions[DownExcursionsCount] = excursion.excursion_size;
                DownExcursionsCount++;
            }
        }
    }

    // Display statistics if enabled.
    if (ShowStatistics)
    {
        DisplayExcursionStatistics();
    }
}

//+------------------------------------------------------------------+
//| Find and mark excursion between two points.                      |
//+------------------------------------------------------------------+
ExcursionInfo FindAndMarkExcursion(int start_idx, int end_idx, bool from_cross_up)
{
    ExcursionInfo info;
    info.cross_time = 0;
    info.excursion_size = 0;
    info.was_up_excursion = from_cross_up;
    info.excursion_time = 0;

    double max_excursion = 0;
    int max_excursion_idx = -1;
    double max_price = 0;

    // Get the cross point value for distance calculation.
    double cross_value = ZigZagBuffer[start_idx];

    // Ensure we process from newer to older bar.
    int from = MathMin(start_idx, end_idx);
    int to = MathMax(start_idx, end_idx);

    if (UseAbsoluteMode == Absolute)
    {
        // Absolute mode: find maximum distance at any point.
        if (from_cross_up)
        {
            // Price crossed above MA - look for highest point.
            double highest_price = 0;
            int highest_idx = -1;
            double reference_value = (UseMAValueAtCross == MA_Value) ? DBL_MAX : cross_value;

            for (int k = from; k <= to; k++)
            {
                double high_k = iHigh(_Symbol, _Period, k);
                if (high_k > highest_price)
                {
                    highest_price = high_k;
                    highest_idx = k;
                }
                if (UseMAValueAtCross == MA_Value && MABuffer[k] < reference_value)
                {
                    reference_value = MABuffer[k];
                }
            }

            if (highest_idx >= 0)
            {
                if (UseMAValueAtCross == MA_Value)
                {
                    max_excursion = highest_price - reference_value;
                }
                else
                {
                    max_excursion = highest_price - cross_value;
                }
                max_excursion_idx = highest_idx;
                max_price = highest_price;
            }

        }
        else
        {
            // Price crossed below MA - look for lowest point.
            double lowest_price = DBL_MAX;
            int lowest_idx = -1;
            double reference_value = (UseMAValueAtCross == MA_Value) ? 0 : cross_value;

            for (int k = from; k <= to; k++)
            {
                double low_k = iLow(_Symbol, _Period, k);
                if (low_k < lowest_price)
                {
                    lowest_price = low_k;
                    lowest_idx = k;
                }
                if (UseMAValueAtCross == MA_Value && MABuffer[k] > reference_value)
                {
                    reference_value = MABuffer[k];
                }
            }

            if (lowest_idx >= 0)
            {
                if (UseMAValueAtCross == MA_Value)
                {
                    max_excursion = reference_value - lowest_price;
                }
                else
                {
                    max_excursion = cross_value - lowest_price;
                }
                max_excursion_idx = lowest_idx;
                max_price = lowest_price;
            }
        }

    }
    else
    {
        // Relative mode: find maximum distance at the same bar.
        for (int k = from; k <= to; k++)
        {
            double distance = 0;
            double price_extreme = 0;

            if (from_cross_up)
            {
                // Price above MA - use high.
                double high_k = iHigh(_Symbol, _Period, k);
                if (UseMAValueAtCross == MA_Value)
                {
                    distance = high_k - MABuffer[k];
                }
                else
                {
                    distance = high_k - cross_value;
                }
                price_extreme = high_k;
            }
            else
            {
                // Price below MA - use low.
                double low_k = iLow(_Symbol, _Period, k);
                if (UseMAValueAtCross == MA_Value)
                {
                    distance = MABuffer[k] - low_k;
                }
                else
                {
                    distance = cross_value - low_k;
                }
                price_extreme = low_k;
            }

            if (distance > max_excursion)
            {
                max_excursion = distance;
                max_excursion_idx = k;
                max_price = price_extreme;
            }
        }
    }

    // Mark the excursion point in ZigZag buffer.
    if (max_excursion_idx >= 0 && max_excursion > 0)
    {
        // Only add to buffer if it's not the start cross point (start cross already has a value).
        // But we CAN add to buffer if it's the end cross point.
        if (max_excursion_idx != start_idx)
        {
            ZigZagBuffer[max_excursion_idx] = max_price;
        }

        // Add text label showing distance (for any bar, including cross bars).
        string label_suffix = from_cross_up ? "_u" : "_d";
        string label_name = "MA_Excursion_" + TimeToString(iTime(_Symbol, _Period, max_excursion_idx), TIME_DATE | TIME_MINUTES) + label_suffix;
        string distance_text = DoubleToString(max_excursion / _Point, 0);

        // Create or update the label
        ObjectDelete(0, label_name);  // Delete if exists (faster than checking)
        ObjectCreate(0, label_name, OBJ_TEXT, 0, iTime(_Symbol, _Period, max_excursion_idx), max_price);
        ObjectSetInteger(0, label_name, OBJPROP_COLOR, from_cross_up ? LabelColorAbove : LabelColorBelow);
        ObjectSetInteger(0, label_name, OBJPROP_FONTSIZE, FontSize);
        ObjectSetInteger(0, label_name, OBJPROP_ANCHOR, from_cross_up ? ANCHOR_LOWER : ANCHOR_UPPER);
        ObjectSetInteger(0, label_name, OBJPROP_SELECTABLE, false);
        ObjectSetInteger(0, label_name, OBJPROP_SELECTED, false);
        ObjectSetInteger(0, label_name, OBJPROP_HIDDEN, true);
        ObjectSetString(0, label_name, OBJPROP_TEXT, distance_text);

        // Set the appropriate excursion buffer value.
        if (from_cross_up)
        {
            UpExcursionBuffer[max_excursion_idx] = max_excursion / _Point;
        }
        else
        {
            DownExcursionBuffer[max_excursion_idx] = max_excursion / _Point;
        }

        // Fill excursion info for alerts.
        info.cross_time = iTime(_Symbol, _Period, start_idx);
        info.excursion_size = max_excursion / _Point;
        info.excursion_time = iTime(_Symbol, _Period, max_excursion_idx);
    }

    return info;
}

//+------------------------------------------------------------------+
//| Display excursion statistics in one of the chart's corner.       |
//+------------------------------------------------------------------+
void DisplayExcursionStatistics()
{
    // Calculate statistics for each type:
    double avgAll = 0, medianAll = 0;
    double avgUp = 0, medianUp = 0;
    double avgDown = 0, medianDown = 0;

    // Determine how many excursions to use.
    int countAll = (StatsCount == 0 || StatsCount >= AllExcursionsCount) ? AllExcursionsCount : StatsCount;
    int countUp = (StatsCount == 0 || StatsCount >= UpExcursionsCount) ? UpExcursionsCount : StatsCount;
    int countDown = (StatsCount == 0 || StatsCount >= DownExcursionsCount) ? DownExcursionsCount : StatsCount;

    // Calculate for all excursions.
    if (countAll > 0)
    {
        double tempAll[];
        ArrayResize(tempAll, countAll);

        // Copy most recent excursions.
        for (int i = 0; i < countAll; i++)
        {
            tempAll[i] = AllExcursions[AllExcursionsCount - countAll + i];
        }

        avgAll = CalculateAverage(tempAll, countAll);
        medianAll = CalculateMedian(tempAll, countAll);
    }

    // Calculate for up excursions.
    if (countUp > 0 && UpExcursionsCount > 0)
    {
        double tempUp[];
        int actualCountUp = MathMin(countUp, UpExcursionsCount);
        ArrayResize(tempUp, actualCountUp);

        // Copy most recent up excursions.
        for (int i = 0; i < actualCountUp; i++)
        {
            tempUp[i] = UpExcursions[UpExcursionsCount - actualCountUp + i];
        }

        avgUp = CalculateAverage(tempUp, actualCountUp);
        medianUp = CalculateMedian(tempUp, actualCountUp);
    }

    // Calculate for down excursions.
    if (countDown > 0 && DownExcursionsCount > 0)
    {
        double tempDown[];
        int actualCountDown = MathMin(countDown, DownExcursionsCount);
        ArrayResize(tempDown, actualCountDown);

        // Copy most recent down excursions.
        for (int i = 0; i < actualCountDown; i++)
        {
            tempDown[i] = DownExcursions[DownExcursionsCount - actualCountDown + i];
        }

        avgDown = CalculateAverage(tempDown, actualCountDown);
        medianDown = CalculateMedian(tempDown, actualCountDown);
    }

    // Prepare display text:
    string header = "MA EXCURSION STATISTICS";
    string countText;

    // Show actual counts when using "All" mode.
    if (StatsCount == 0)
    {
        countText = "All " + IntegerToString(AllExcursionsCount) + " Excursions";
    }
    else
    {
        int actualCount = MathMin(StatsCount, AllExcursionsCount);
        countText = IntegerToString(actualCount) + " Recent Excursions";
    }

    // Create labels:
    int yOffset = 20;
    int xOffset = 10;

    // Adjust X offset for right-side corners.
    if (StatsCorner == CORNER_RIGHT_UPPER || StatsCorner == CORNER_RIGHT_LOWER)
    {
        xOffset = 180;  // Larger offset for right-side corners.
    }

    // Apply user-defined offset.
    xOffset += StatsXOffset;

    // Check if we need to invert order for lower corners.
    bool invertOrder = (StatsCorner == CORNER_LEFT_LOWER || StatsCorner == CORNER_RIGHT_LOWER);

    if (invertOrder)
    {
        // For lower corners, display from bottom to top.
        CreateCornerLabel("Stats_Down_Med", "  Med: " + DoubleToString(medianDown, 0) + " points", xOffset, yOffset, LabelColorBelow, StatsFontSize, false);
        yOffset += 15;
        CreateCornerLabel("Stats_Down_Avg", "  Avg: " + DoubleToString(avgDown, 0) + " points", xOffset, yOffset, LabelColorBelow, StatsFontSize, false);
        yOffset += 15;
        CreateCornerLabel("Stats_Down", "DOWN:", xOffset, yOffset, LabelColorBelow, StatsFontSize, true);
        yOffset += 20;

        CreateCornerLabel("Stats_Up_Med", "  Med: " + DoubleToString(medianUp, 0) + " points", xOffset, yOffset, LabelColorAbove, StatsFontSize, false);
        yOffset += 15;
        CreateCornerLabel("Stats_Up_Avg", "  Avg: " + DoubleToString(avgUp, 0) + " points", xOffset, yOffset, LabelColorAbove, StatsFontSize, false);
        yOffset += 15;
        CreateCornerLabel("Stats_Up", "UP:", xOffset, yOffset, LabelColorAbove, StatsFontSize, true);
        yOffset += 20;

        CreateCornerLabel("Stats_Total_Med", "  Med: " + DoubleToString(medianAll, 0) + " points", xOffset, yOffset, StatsColor, StatsFontSize, false);
        yOffset += 15;
        CreateCornerLabel("Stats_Total_Avg", "  Avg: " + DoubleToString(avgAll, 0) + " points", xOffset, yOffset, StatsColor, StatsFontSize, false);
        yOffset += 15;
        CreateCornerLabel("Stats_Total", "TOTAL:", xOffset, yOffset, StatsColor, StatsFontSize, true);
        yOffset += 25;

        CreateCornerLabel("Stats_Count", "(" + countText + ")", xOffset, yOffset, StatsColor, StatsFontSize - 1, false);
        yOffset += 20;
        CreateCornerLabel("Stats_Header", header, xOffset, yOffset, StatsColor, StatsFontSize + 1, true);
    }
    else
    {
        // For upper corners, display from top to bottom (normal order).
        CreateCornerLabel("Stats_Header", header, xOffset, yOffset, StatsColor, StatsFontSize + 1, true);
        yOffset += 20;

        CreateCornerLabel("Stats_Count", "(" + countText + ")", xOffset, yOffset, StatsColor, StatsFontSize - 1, false);
        yOffset += 25;

        CreateCornerLabel("Stats_Total", "TOTAL:", xOffset, yOffset, StatsColor, StatsFontSize, true);
        yOffset += 15;
        CreateCornerLabel("Stats_Total_Avg", "  Avg: " + DoubleToString(avgAll, 0) + " points", xOffset, yOffset, StatsColor, StatsFontSize, false);
        yOffset += 15;
        CreateCornerLabel("Stats_Total_Med", "  Med: " + DoubleToString(medianAll, 0) + " points", xOffset, yOffset, StatsColor, StatsFontSize, false);
        yOffset += 20;

        CreateCornerLabel("Stats_Up", "UP:", xOffset, yOffset, LabelColorAbove, StatsFontSize, true);
        yOffset += 15;
        CreateCornerLabel("Stats_Up_Avg", "  Avg: " + DoubleToString(avgUp, 0) + " points", xOffset, yOffset, LabelColorAbove, StatsFontSize, false);
        yOffset += 15;
        CreateCornerLabel("Stats_Up_Med", "  Med: " + DoubleToString(medianUp, 0) + " points", xOffset, yOffset, LabelColorAbove, StatsFontSize, false);
        yOffset += 20;

        CreateCornerLabel("Stats_Down", "DOWN:", xOffset, yOffset, LabelColorBelow, StatsFontSize, true);
        yOffset += 15;
        CreateCornerLabel("Stats_Down_Avg", "  Avg: " + DoubleToString(avgDown, 0) + " points", xOffset, yOffset, LabelColorBelow, StatsFontSize, false);
        yOffset += 15;
        CreateCornerLabel("Stats_Down_Med", "  Med: " + DoubleToString(medianDown, 0) + " points", xOffset, yOffset, LabelColorBelow, StatsFontSize, false);
    }
}

//+------------------------------------------------------------------+
//| Create corner label.                                             |
//+------------------------------------------------------------------+
void CreateCornerLabel(string name, string text, int x, int y, color clr, int fontSize, bool bold)
{
    string fullName = "MA_Excursion_" + name;

    // Always delete and recreate (faster than checking existence).
    ObjectDelete(0, fullName);
    ObjectCreate(0, fullName, OBJ_LABEL, 0, 0, 0);
    ObjectSetInteger(0, fullName, OBJPROP_CORNER, StatsCorner);
    ObjectSetInteger(0, fullName, OBJPROP_XDISTANCE, x);
    ObjectSetInteger(0, fullName, OBJPROP_YDISTANCE, y);
    ObjectSetInteger(0, fullName, OBJPROP_COLOR, clr);
    ObjectSetInteger(0, fullName, OBJPROP_FONTSIZE, fontSize);
    ObjectSetString(0, fullName, OBJPROP_FONT, bold ? "Arial Bold" : "Arial");
    ObjectSetInteger(0, fullName, OBJPROP_SELECTABLE, false);
    ObjectSetInteger(0, fullName, OBJPROP_SELECTED, false);
    ObjectSetInteger(0, fullName, OBJPROP_HIDDEN, true);
    ObjectSetString(0, fullName, OBJPROP_TEXT, text);
}

//+------------------------------------------------------------------+
//| Calculate average of array.                                      |
//+------------------------------------------------------------------+
double CalculateAverage(double &arr[], int count)
{
    if (count == 0) return 0;

    double sum = 0;
    for (int i = 0; i < count; i++)
    {
        sum += arr[i];
    }

    return sum / count;
}

//+------------------------------------------------------------------+
//| Calculate median of array.                                       |
//+------------------------------------------------------------------+
double CalculateMedian(double &arr[], int count)
{
    if (count == 0) return 0;

    // Sort array:
    double sorted[];
    ArrayResize(sorted, count);
    ArrayCopy(sorted, arr);
    ArraySort(sorted);

    // Calculate median:
    if (count % 2 == 0)
    {
        // Even number of elements.
        return (sorted[count / 2 - 1] + sorted[count / 2]) / 2.0;
    }
    else
    {
        // Odd number of elements.
        return sorted[count / 2];
    }
}

//+------------------------------------------------------------------+
//| Send alert about completed excursion.                            |
//+------------------------------------------------------------------+
void SendExcursionAlert(ExcursionInfo &info)
{
    // Prepare alert message:
    string direction = info.was_up_excursion ? "UP" : "DOWN";
    
    // Message for popup alerts doesn't need indicator/symbol/TF info.
    string message_alert = StringFormat("Previous cross: %s | %s excursion: %.0f points",
                                  TimeToString(info.cross_time, TIME_DATE | TIME_MINUTES),
                                  direction,
                                  info.excursion_size);

    // Email and push-notification messages need that info.
    string message = StringFormat("MA Cross Excursion: %s %s | %s",
                                  _Symbol,
                                  PeriodToString(_Period),
                                  message_alert);
    
    // Send alerts based on settings.
    if (UsePopupAlert)
    {
        Alert(message);
    }
    if (UseEmailAlert)
    {
        SendMail("MA Cross Excursion Alert", message);
    }
    if (UsePushAlert)
    {
        SendNotification(message);
    }
    if (UseSoundAlert)
    {
        PlaySound(SoundFile);
    }
}

//+------------------------------------------------------------------+
//| Convert period to string.                                        |
//+------------------------------------------------------------------+
string PeriodToString(ENUM_TIMEFRAMES period)
{
    switch(period)
    {
    case PERIOD_M1:
        return "M1";
    case PERIOD_M2:
        return "M2";
    case PERIOD_M3:
        return "M3";
    case PERIOD_M4:
        return "M4";
    case PERIOD_M5:
        return "M5";
    case PERIOD_M6:
        return "M6";
    case PERIOD_M10:
        return "M10";
    case PERIOD_M12:
        return "M12";
    case PERIOD_M15:
        return "M15";
    case PERIOD_M20:
        return "M20";
    case PERIOD_M30:
        return "M30";
    case PERIOD_H1:
        return "H1";
    case PERIOD_H2:
        return "H2";
    case PERIOD_H3:
        return "H3";
    case PERIOD_H4:
        return "H4";
    case PERIOD_H6:
        return "H6";
    case PERIOD_H8:
        return "H8";
    case PERIOD_H12:
        return "H12";
    case PERIOD_D1:
        return "D1";
    case PERIOD_W1:
        return "W1";
    case PERIOD_MN1:
        return "MN1";
    default:
        return "Unknown";
    }
}

//+------------------------------------------------------------------+
//| Create on/off button.                                            |
//+------------------------------------------------------------------+
void CreateButton()
{
    ObjectCreate(0, ButtonName, OBJ_BUTTON, 0, 0, 0);
    ObjectSetInteger(0, ButtonName, OBJPROP_CORNER, ButtonCorner);
    ObjectSetInteger(0, ButtonName, OBJPROP_XDISTANCE, ButtonXOffset);
    ObjectSetInteger(0, ButtonName, OBJPROP_YDISTANCE, ButtonYOffset);
    ObjectSetInteger(0, ButtonName, OBJPROP_XSIZE, 80);
    ObjectSetInteger(0, ButtonName, OBJPROP_YSIZE, 25);
    ObjectSetInteger(0, ButtonName, OBJPROP_FONTSIZE, 9);
    ObjectSetString(0, ButtonName, OBJPROP_FONT, "Arial");
    ObjectSetInteger(0, ButtonName, OBJPROP_STATE, false);
    ObjectSetInteger(0, ButtonName, OBJPROP_SELECTABLE, false);
    ObjectSetInteger(0, ButtonName, OBJPROP_SELECTED, false);
    ObjectSetInteger(0, ButtonName, OBJPROP_HIDDEN, true);

    // Set initial appearance.
    UpdateButton();
}

//+------------------------------------------------------------------+
//| Update button appearance based on indicator state.               |
//+------------------------------------------------------------------+
void UpdateButton()
{
    if (IndicatorEnabled)
    {
        ObjectSetString(0, ButtonName, OBJPROP_TEXT, "MA Exc: ON");
        ObjectSetInteger(0, ButtonName, OBJPROP_BGCOLOR, clrGreen);
        ObjectSetInteger(0, ButtonName, OBJPROP_COLOR, clrWhite);
    }
    else
    {
        ObjectSetString(0, ButtonName, OBJPROP_TEXT, "MA Exc: OFF");
        ObjectSetInteger(0, ButtonName, OBJPROP_BGCOLOR, clrRed);
        ObjectSetInteger(0, ButtonName, OBJPROP_COLOR, clrWhite);
    }
    ChartRedraw();
}
//+------------------------------------------------------------------+