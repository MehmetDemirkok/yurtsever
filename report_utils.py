import pandas as pd
from datetime import datetime
from typing import List, Dict, Any
from helpers import get_date_range, format_currency

class StayReport:
    def __init__(self, stays: List[Dict[str, Any]]):
        """Initialize stay report with stay data.
        
        Args:
            stays (List[Dict[str, Any]]): List of stay records
        """
        self.stays = stays
        self.df = pd.DataFrame(stays)
        if not self.df.empty:
            self.df['check_in_date'] = pd.to_datetime(self.df['check_in_date'])
            self.df['check_out_date'] = pd.to_datetime(self.df['check_out_date'])
    
    def get_period_report(self, period: str) -> pd.DataFrame:
        """Generate report for a specific period.
        
        Args:
            period (str): One of 'week', 'month', 'year'
            
        Returns:
            pd.DataFrame: Filtered and formatted report
        """
        if self.df.empty:
            return pd.DataFrame()
            
        start_date, end_date = get_date_range(period)
        
        # Filter stays within the period
        mask = (
            (self.df['check_in_date'] >= start_date) & 
            (self.df['check_in_date'] <= end_date)
        )
        period_df = self.df[mask].copy()
        
        if period_df.empty:
            return pd.DataFrame()
        
        # Calculate summary statistics
        summary = {
            'total_stays': len(period_df),
            'total_revenue': period_df['total_amount'].sum(),
            'average_stay_duration': (
                (period_df['check_out_date'] - period_df['check_in_date']).dt.days.mean()
            ),
            'average_nightly_rate': period_df['nightly_rate'].mean()
        }
        
        # Format the report
        report_df = period_df[[
            'guest_name', 'room_type', 'check_in_date', 
            'check_out_date', 'nightly_rate', 'total_amount'
        ]].copy()
        
        report_df['check_in_date'] = report_df['check_in_date'].dt.strftime('%Y-%m-%d')
        report_df['check_out_date'] = report_df['check_out_date'].dt.strftime('%Y-%m-%d')
        report_df['nightly_rate'] = report_df['nightly_rate'].apply(format_currency)
        report_df['total_amount'] = report_df['total_amount'].apply(format_currency)
        
        # Add summary row
        summary_row = pd.DataFrame([{
            'guest_name': 'SUMMARY',
            'room_type': '',
            'check_in_date': '',
            'check_out_date': '',
            'nightly_rate': format_currency(summary['average_nightly_rate']),
            'total_amount': format_currency(summary['total_revenue'])
        }])
        
        report_df = pd.concat([report_df, summary_row], ignore_index=True)
        
        return report_df
    
    def get_room_occupancy(self, period: str) -> pd.DataFrame:
        """Generate room occupancy report for a specific period.
        
        Args:
            period (str): One of 'week', 'month', 'year'
            
        Returns:
            pd.DataFrame: Room occupancy report
        """
        if self.df.empty:
            return pd.DataFrame()
            
        start_date, end_date = get_date_range(period)
        
        # Filter stays within the period
        mask = (
            (self.df['check_in_date'] >= start_date) & 
            (self.df['check_in_date'] <= end_date)
        )
        period_df = self.df[mask].copy()
        
        if period_df.empty:
            return pd.DataFrame()
        
        # Group by room type and calculate statistics
        occupancy_df = period_df.groupby('room_type').agg({
            'id': 'count',
            'total_amount': 'sum',
            'nightly_rate': 'mean'
        }).reset_index()
        
        occupancy_df.columns = ['room_type', 'number_of_stays', 'total_revenue', 'average_rate']
        
        # Format the report
        occupancy_df['total_revenue'] = occupancy_df['total_revenue'].apply(format_currency)
        occupancy_df['average_rate'] = occupancy_df['average_rate'].apply(format_currency)
        
        return occupancy_df.sort_values('number_of_stays', ascending=False) 