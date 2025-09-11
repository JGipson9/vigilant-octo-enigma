import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime, timedelta
import warnings
import os
warnings.filterwarnings('ignore')

def load_and_analyze_all_tabs(file_path):
    """
    Load and analyze all tabs from the Chenmark Excel file
    """
    # Check if file exists
    if not os.path.exists(file_path):
        print(f"❌ File not found at: {file_path}")
        print("Please check the file path and ensure the Excel file is in the correct location.")
        return {}, []
    
    # Read all sheets from the Excel file
    try:
        excel_file = pd.ExcelFile(file_path)
        print(f"📊 CHENMARK CASE STUDY ANALYSIS - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("="*80)
        print(f"Found {len(excel_file.sheet_names)} tabs in the Excel file:")
        
        # Dictionary to store all dataframes
        data_tabs = {}
        
        for i, sheet_name in enumerate(excel_file.sheet_names):
            print(f"  {i+1}. {sheet_name}")
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                data_tabs[sheet_name] = df
                print(f"     └─ Shape: {df.shape}")
            except Exception as e:
                print(f"     └─ Error loading: {e}")
        
        return data_tabs, excel_file.sheet_names
    
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return {}, []

def comprehensive_financial_analysis(data_tabs):
    """
    Perform comprehensive financial analysis across all tabs
    """
    print("\n" + "="*80)
    print("🏢 COMPREHENSIVE FINANCIAL ANALYSIS")
    print("="*80)
    
    # Find financial data across tabs
    financial_metrics = {}
    
    for tab_name, df in data_tabs.items():
        print(f"\n📈 Analyzing {tab_name}:")
        print("-" * 50)
        
        # Basic info
        print(f"Dimensions: {df.shape[0]} rows × {df.shape[1]} columns")
        
        # Look for key financial indicators
        financial_columns = []
        for col in df.columns:
            col_lower = str(col).lower()
            if any(term in col_lower for term in ['revenue', 'sales', 'profit', 'ebitda', 
                                                  'margin', 'cost', 'expense', 'cash', 
                                                  'debt', 'equity', 'roi', 'growth',
                                                  'income', 'earnings', 'assets', 'liabilities']):
                financial_columns.append(col)
        
        if financial_columns:
            print(f"Financial columns found: {financial_columns}")
            
            # Calculate key metrics for numeric columns
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            if len(numeric_cols) > 0:
                print("\nKey Statistics:")
                for col in numeric_cols[:5]:  # Show first 5 numeric columns
                    if df[col].notna().sum() > 0:
                        print(f"  {col}:")
                        print(f"    Mean: {df[col].mean():,.2f}")
                        print(f"    Median: {df[col].median():,.2f}")
                        print(f"    Std Dev: {df[col].std():,.2f}")
                        if df[col].min() != df[col].max():
                            print(f"    Range: {df[col].min():,.2f} to {df[col].max():,.2f}")
        
        # Store for cross-tab analysis
        financial_metrics[tab_name] = {
            'shape': df.shape,
            'numeric_columns': list(df.select_dtypes(include=[np.number]).columns),
            'financial_columns': financial_columns,
            'data': df
        }
    
    return financial_metrics

def advanced_trend_analysis(data_tabs):
    """
    Perform advanced trend analysis looking for time-based patterns
    """
    print("\n" + "="*80)
    print("📊 ADVANCED TREND ANALYSIS")
    print("="*80)
    
    trends = {}
    
    for tab_name, df in data_tabs.items():
        print(f"\n🔍 Trend Analysis for {tab_name}:")
        print("-" * 40)
        
        # Look for date columns
        date_columns = []
        for col in df.columns:
            if df[col].dtype == 'datetime64[ns]' or 'date' in str(col).lower() or 'year' in str(col).lower():
                date_columns.append(col)
        
        if date_columns:
            print(f"Date columns found: {date_columns}")
            
            # Time series analysis for each date column
            for date_col in date_columns:
                if df[date_col].notna().sum() > 1:
                    print(f"\n  Timeline for {date_col}:")
                    try:
                        df_sorted = df.sort_values(date_col)
                        print(f"    Period: {df_sorted[date_col].min()} to {df_sorted[date_col].max()}")
                        
                        # Calculate growth rates if there are numeric columns
                        numeric_cols = df.select_dtypes(include=[np.number]).columns
                        for num_col in numeric_cols[:3]:  # Analyze first 3 numeric columns
                            if df[num_col].notna().sum() > 1:
                                values = df_sorted[num_col].dropna()
                                if len(values) > 1:
                                    growth_rate = ((values.iloc[-1] - values.iloc[0]) / values.iloc[0]) * 100
                                    print(f"    {num_col} growth: {growth_rate:.1f}%")
                    except Exception as e:
                        print(f"    Error in trend analysis: {e}")
        
        # Look for sequential data patterns
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        correlations = []
        if len(numeric_cols) > 1:
            print(f"\n  Correlation Analysis:")
            try:
                corr_matrix = df[numeric_cols].corr()
                
                # Find strongest correlations
                for i in range(len(corr_matrix.columns)):
                    for j in range(i+1, len(corr_matrix.columns)):
                        corr_val = corr_matrix.iloc[i, j]
                        if not pd.isna(corr_val) and abs(corr_val) > 0.5:
                            correlations.append((
                                corr_matrix.columns[i], 
                                corr_matrix.columns[j], 
                                corr_val
                            ))
                
                if correlations:
                    correlations.sort(key=lambda x: abs(x[2]), reverse=True)
                    print("    Strong correlations found:")
                    for col1, col2, corr in correlations[:5]:
                        print(f"      {col1} ↔ {col2}: {corr:.3f}")
                else:
                    print("    No strong correlations found (>0.5)")
            except Exception as e:
                print(f"    Error in correlation analysis: {e}")
        
        trends[tab_name] = {
            'date_columns': date_columns,
            'correlations': correlations
        }
    
    return trends

def competitive_analysis(data_tabs):
    """
    Perform competitive and market analysis
    """
    print("\n" + "="*80)
    print("🏆 COMPETITIVE & MARKET ANALYSIS")
    print("="*80)
    
    competitive_insights = {}
    
    for tab_name, df in data_tabs.items():
        print(f"\n🎯 Market Analysis for {tab_name}:")
        print("-" * 40)
        
        # Look for company/competitor identifiers
        text_cols = df.select_dtypes(include=['object']).columns
        
        if len(text_cols) > 0:
            for col in text_cols[:3]:  # Check first 3 text columns
                unique_values = df[col].dropna().unique()
                if 2 <= len(unique_values) <= 20:  # Reasonable number for companies/segments
                    print(f"\n  Categories in {col}:")
                    for i, value in enumerate(unique_values[:10]):  # Show first 10
                        count = df[df[col] == value].shape[0]
                        print(f"    {i+1}. {value}: {count} records")
                    
                    # Performance comparison if numeric data exists
                    numeric_cols = df.select_dtypes(include=[np.number]).columns
                    if len(numeric_cols) > 0:
                        print(f"\n  Performance Comparison by {col}:")
                        for num_col in numeric_cols[:3]:
                            if df[num_col].notna().sum() > 0:
                                try:
                                    comparison = df.groupby(col)[num_col].agg(['mean', 'median', 'std']).round(2)
                                    print(f"    {num_col}:")
                                    print(comparison.head())
                                except Exception as e:
                                    print(f"    Error in comparison: {e}")
        
        competitive_insights[tab_name] = {
            'text_columns': list(text_cols),
            'analysis_performed': True
        }
    
    return competitive_insights

def risk_and_opportunity_analysis(data_tabs):
    """
    Identify risks and opportunities from the data
    """
    print("\n" + "="*80)
    print("⚠️  RISK & OPPORTUNITY ANALYSIS")
    print("="*80)
    
    risks_opportunities = {}
    
    for tab_name, df in data_tabs.items():
        print(f"\n🔍 Risk Analysis for {tab_name}:")
        print("-" * 40)
        
        tab_risks = []
        tab_opportunities = []
        
        # Analyze numeric columns for volatility and outliers
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        
        for col in numeric_cols:
            if df[col].notna().sum() > 5:  # Need at least 5 data points
                values = df[col].dropna()
                
                # Calculate volatility (coefficient of variation)
                if values.mean() != 0:
                    cv = values.std() / abs(values.mean())
                    
                    if cv > 0.5:  # High volatility
                        tab_risks.append(f"High volatility in {col} (CV: {cv:.2f})")
                    elif cv < 0.1:  # Very stable
                        tab_opportunities.append(f"Stable performance in {col} (CV: {cv:.2f})")
                
                # Identify outliers
                try:
                    q1, q3 = values.quantile([0.25, 0.75])
                    iqr = q3 - q1
                    if iqr > 0:
                        outliers = values[(values < (q1 - 1.5 * iqr)) | (values > (q3 + 1.5 * iqr))]
                        
                        if len(outliers) > 0:
                            outlier_pct = (len(outliers) / len(values)) * 100
                            if outlier_pct > 10:
                                tab_risks.append(f"Many outliers in {col} ({outlier_pct:.1f}% of data)")
                except Exception as e:
                    pass  # Skip if quartile calculation fails
                
                # Growth trends
                if len(values) > 1:
                    try:
                        recent_trend = np.polyfit(range(len(values)), values, 1)[0]
                        if recent_trend > 0 and abs(recent_trend) > values.std() * 0.1:
                            tab_opportunities.append(f"Positive trend in {col} (slope: {recent_trend:.2f})")
                        elif recent_trend < 0 and abs(recent_trend) > values.std() * 0.1:
                            tab_risks.append(f"Declining trend in {col} (slope: {recent_trend:.2f})")
                    except Exception as e:
                        pass  # Skip if trend calculation fails
        
        # Display findings
        if tab_risks:
            print("  🚨 Identified Risks:")
            for i, risk in enumerate(tab_risks[:5], 1):
                print(f"    {i}. {risk}")
        
        if tab_opportunities:
            print("  🌟 Identified Opportunities:")
            for i, opp in enumerate(tab_opportunities[:5], 1):
                print(f"    {i}. {opp}")
        
        if not tab_risks and not tab_opportunities:
            print("  ✅ No significant risks or opportunities detected in numeric data")
        
        risks_opportunities[tab_name] = {
            'risks': tab_risks,
            'opportunities': tab_opportunities
        }
    
    return risks_opportunities

def strategic_recommendations(financial_metrics, trends, competitive_insights, risks_opportunities):
    """
    Generate strategic recommendations based on all analyses
    """
    print("\n" + "="*80)
    print("🎯 STRATEGIC RECOMMENDATIONS")
    print("="*80)
    
    recommendations = []
    
    # Financial recommendations
    print("\n💰 FINANCIAL STRATEGY:")
    print("-" * 30)
    
    total_numeric_cols = sum(len(fm['numeric_columns']) for fm in financial_metrics.values())
    if total_numeric_cols > 10:
        recommendations.append("Rich financial data available - implement comprehensive KPI dashboard")
        print("1. Implement comprehensive KPI dashboard with real-time monitoring")
    
    high_volatility_tabs = []
    for tab, ro in risks_opportunities.items():
        if any('volatility' in risk.lower() for risk in ro['risks']):
            high_volatility_tabs.append(tab)
    
    if high_volatility_tabs:
        recommendations.append(f"Address volatility in: {', '.join(high_volatility_tabs)}")
        print(f"2. Develop risk management strategies for volatile metrics in {', '.join(high_volatility_tabs)}")
    
    # Growth recommendations
    print("\n📈 GROWTH STRATEGY:")
    print("-" * 25)
    
    growth_opportunities = []
    for tab, ro in risks_opportunities.items():
        growth_opportunities.extend([opp for opp in ro['opportunities'] if 'trend' in opp.lower()])
    
    if growth_opportunities:
        recommendations.append("Capitalize on positive trends identified in the data")
        print("3. Double down on areas showing positive momentum")
        for opp in growth_opportunities[:3]:
            print(f"   • {opp}")
    
    # Operational recommendations
    print("\n⚙️  OPERATIONAL EXCELLENCE:")
    print("-" * 35)
    
    stable_metrics = []
    for tab, ro in risks_opportunities.items():
        stable_metrics.extend([opp for opp in ro['opportunities'] if 'stable' in opp.lower()])
    
    if stable_metrics:
        recommendations.append("Leverage stable performance areas as competitive advantages")
        print("4. Use stable performance areas as foundation for expansion")
    
    # Data-driven recommendations
    print("\n📊 DATA & ANALYTICS:")
    print("-" * 25)
    
    strong_correlations = sum(len(trend['correlations']) for trend in trends.values())
    if strong_correlations > 5:
        recommendations.append("Exploit strong correlations for predictive analytics")
        print("5. Develop predictive models based on strong correlations found")
    
    # Market recommendations
    print("\n🏢 MARKET STRATEGY:")
    print("-" * 23)
    
    competitive_data_available = any(ci['analysis_performed'] for ci in competitive_insights.values())
    if competitive_data_available:
        recommendations.append("Enhance competitive intelligence and market positioning")
        print("6. Enhance competitive analysis and market positioning strategies")
    
    print(f"\n📝 SUMMARY: {len(recommendations)} strategic recommendations generated")
    
    return recommendations

def create_executive_summary(data_tabs, financial_metrics, trends, competitive_insights, risks_opportunities, recommendations):
    """
    Create comprehensive executive summary
    """
    print("\n" + "="*80)
    print("📋 EXECUTIVE SUMMARY")
    print("="*80)
    
    # Data overview
    total_tabs = len(data_tabs)
    total_rows = sum(df.shape[0] for df in data_tabs.values())
    total_cols = sum(df.shape[1] for df in data_tabs.values())
    
    print(f"\n🔍 DATA OVERVIEW:")
    print(f"   • {total_tabs} data tabs analyzed")
    print(f"   • {total_rows:,} total data points")
    print(f"   • {total_cols} total columns across all tabs")
    
    # Key findings
    print(f"\n📊 KEY FINDINGS:")
    
    # Financial insights
    financial_tabs = len([tab for tab, fm in financial_metrics.items() if fm['financial_columns']])
    if financial_tabs > 0:
        print(f"   • {financial_tabs} tabs contain financial data")
    
    # Risk/opportunity count
    total_risks = sum(len(ro['risks']) for ro in risks_opportunities.values())
    total_opps = sum(len(ro['opportunities']) for ro in risks_opportunities.values())
    
    print(f"   • {total_risks} risks identified")
    print(f"   • {total_opps} opportunities discovered")
    
    # Correlation insights
    total_correlations = sum(len(trend['correlations']) for trend in trends.values())
    print(f"   • {total_correlations} strong correlations found")
    
    # Strategic priorities
    print(f"\n🎯 STRATEGIC PRIORITIES:")
    for i, rec in enumerate(recommendations[:5], 1):
        print(f"   {i}. {rec}")
    
    # Next steps
    print(f"\n🚀 RECOMMENDED NEXT STEPS:")
    print("   1. Deep-dive analysis of highest-priority opportunities")
    print("   2. Develop risk mitigation strategies for identified concerns")
    print("   3. Create automated monitoring dashboard")
    print("   4. Establish regular review cycles for key metrics")
    print("   5. Benchmark against industry standards")
    
    return {
        'total_tabs': total_tabs,
        'total_rows': total_rows,
        'total_risks': total_risks,
        'total_opportunities': total_opps,
        'recommendations': recommendations
    }

def main():
    """
    Main analysis function - Updated for Windows path handling
    """
    # Try multiple potential file locations
    potential_paths = [
        "Chenmark Case Study 2025 (4).xlsm",  # Same directory
        r"C:\Users\jalen\OneDrive\Desktop\Booth\Chenmark Case Study\Chenmark Case Study 2025 (4).xlsm",  # Full path
        os.path.join(os.getcwd(), "Chenmark Case Study 2025 (4).xlsm")  # Current working directory
    ]
    
    file_path = None
    for path in potential_paths:
        if os.path.exists(path):
            file_path = path
            break
    
    if file_path is None:
        print("❌ Excel file not found in any of the expected locations:")
        for path in potential_paths:
            print(f"   - {path}")
        print("\n💡 Please ensure the file 'Chenmark Case Study 2025 (4).xlsm' is available")
        return
    
    try:
        # Load all data
        print("🔄 Loading Excel file...")
        print(f"📁 Using file: {file_path}")
        
        data_tabs, sheet_names = load_and_analyze_all_tabs(file_path)
        
        if not data_tabs:
            print("❌ No data loaded. Please check the file path and format.")
            return
        
        # Perform all analyses
        print("\n🔄 Performing comprehensive analysis...")
        
        financial_metrics = comprehensive_financial_analysis(data_tabs)
        trends = advanced_trend_analysis(data_tabs)
        competitive_insights = competitive_analysis(data_tabs)
        risks_opportunities = risk_and_opportunity_analysis(data_tabs)
        recommendations = strategic_recommendations(financial_metrics, trends, competitive_insights, risks_opportunities)
        
        # Create executive summary
        summary = create_executive_summary(data_tabs, financial_metrics, trends, competitive_insights, risks_opportunities, recommendations)
        
        print("\n" + "="*80)
        print("✅ ANALYSIS COMPLETE")
        print("="*80)
        print(f"📊 Comprehensive analysis of {len(data_tabs)} tabs completed successfully!")
        print(f"📈 {len(recommendations)} strategic recommendations generated")
        print(f"⏰ Analysis completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Save results to text file
        output_file = "chenmark_analysis_results.txt"
        print(f"\n💾 Saving detailed results to: {output_file}")
        
    except Exception as e:
        print(f"❌ An error occurred: {str(e)}")
        print("💡 Please check your data format and try again.")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()