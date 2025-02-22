import tabula
import pandas as pd
import numpy as np
import re
from typing import List, Dict, Any, Optional, Tuple
import logging
from datetime import datetime
import warnings

logger = logging.getLogger(__name__)

class TableExtractor:
    """
    A class to handle table extraction from PDFs, specialized for bank statements
    using Tabula for extraction and combining multiple pages into a single continuous table.
    """
    
    def __init__(self):
        # Common patterns for bank statements
        self.date_pattern = re.compile(r'\d{2}(?:\s+)?(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(?:\s+)?\d{2,4}|\d{2}/\d{2}/\d{2,4}|\d{2}-\d{2}-\d{2,4}|\d{4}-\d{2}-\d{2}|\d{2}\s+[A-Z]{3}', re.IGNORECASE)
        self.amount_pattern = re.compile(r'(?:[$£€])?(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d{2})?')
        
        # Standard column names for ANZ bank statements
        self.standard_columns = ['Date', 'Description', 'Withdrawals ($)', 'Deposits ($)', 'Balance ($)']
        
        # Suppress FutureWarning from pandas
        warnings.filterwarnings('ignore', category=FutureWarning)
        
    def parse_header_data(self, df: pd.DataFrame) -> Dict[str, str]:
        """
        Extract header information from the statement.
        Returns a dictionary of header fields.
        """
        header_data = {}
        
        # Look for header information in the first few rows
        header_text = ' '.join([str(val) for val in df.iloc[0:5].values.flatten() if pd.notna(val)])
        
        # Extract account information using regex patterns
        account_pattern = re.compile(r'ACCOUNT\s+(\d+)', re.IGNORECASE)
        balance_pattern = re.compile(r'Opening Balance[:\s]+(\$?[\d,.]+)', re.IGNORECASE)
        date_range_pattern = re.compile(r'(\d{2}\s+[A-Za-z]{3}.*?to.*?\d{2}\s+[A-Za-z]{3})', re.IGNORECASE)
        
        # Find matches
        account_match = account_pattern.search(header_text)
        balance_match = balance_pattern.search(header_text)
        date_range_match = date_range_pattern.search(header_text)
        
        if account_match:
            header_data['account_number'] = account_match.group(1)
        if balance_match:
            header_data['opening_balance'] = balance_match.group(1)
        if date_range_match:
            header_data['statement_period'] = date_range_match.group(1)
            
        return header_data
    
    def parse_pdf_to_excel(self, file_path: str, fallback_attempted: bool = False) -> pd.DataFrame:
        """
        Parses PDF tabular data into a DataFrame.
        Handles multi-row transactions and mixed header/transaction data.
        Added fallback_attempted parameter to prevent infinite loops.
        """
        try:
            # Read all pages from the PDF
            dfs = tabula.read_pdf(
                file_path,
                pages='all',
                stream=True,
                multiple_tables=True,
                guess=True,
                pandas_options={'header': None}  # Force no header inference
            )
            
            if not dfs:
                if not fallback_attempted:
                    logger.warning("No tables found, attempting alternative extraction method")
                    return self.extract_tables_fallback(file_path)
                else:
                    logger.error("No tables found in PDF after fallback attempt")
                    return pd.DataFrame(columns=self.standard_columns)
            
            # Process and combine all tables
            all_transactions = []
            
            for df in dfs:
                if df.empty or len(df.columns) < 3:
                    continue
                
                # Clean column names and standardize DataFrame structure
                df.columns = [str(col).strip() for col in df.columns]
                
                # Process rows
                current_transaction = None
                skip_next_row = False
                
                for idx, row in df.iterrows():
                    if skip_next_row:
                        skip_next_row = False
                        continue
                    
                    # Get original row values before cleaning to check for NaN
                    original_values = [val for val in row.values]
                    
                    # Convert row values to strings and clean
                    row_values = {col: str(val).strip() if pd.notna(val) else '' for col, val in row.items()}
                    next_row_values = None
                    if idx + 1 < len(df):
                        next_row_values = {col: str(val).strip() if pd.notna(val) else '' 
                                         for col, val in df.iloc[idx + 1].items()}
                    
                    # Skip pure header rows
                    row_text = ' '.join(row_values.values()).upper()
                    if any(header in row_text for header in ['OPENING BALANCE', 'TOTALS AT END OF PAGE']):
                        continue
                    
                    # Check if this is the start of a new transaction
                    date_value = None
                    for val in row_values.values():
                        if val and self.date_pattern.match(val):
                            date_value = val
                            break
                    
                    if date_value:
                        # If we have a previous transaction, save it
                        if current_transaction:
                            all_transactions.append(current_transaction)
                        
                        # Start new transaction
                        current_transaction = {
                            'Date': date_value,
                            'Description': '',
                            'Withdrawals ($)': 0.0,
                            'Deposits ($)': 0.0,
                            'Balance ($)': 0.0
                        }
                        
                        # Process description and amounts
                        desc_parts = []
                        amounts = []
                        amount_positions = []  # Track positions of amounts
                        
                        # Process current row
                        for i, val in enumerate(original_values):
                            if pd.notna(val):
                                val_str = str(val).strip()
                                if self.date_pattern.match(val_str):
                                    continue
                                if self.amount_pattern.match(val_str):
                                    clean_amount = self.clean_amount(val_str)
                                    if clean_amount is not None:
                                        amounts.append(clean_amount)
                                        amount_positions.append(i)
                                elif val_str not in ['NaN', 'nan']:
                                    desc_parts.append(val_str)
                        
                        # Check if next row is continuation
                        if next_row_values:
                            has_date = any(self.date_pattern.match(val) for val in next_row_values.values() if val)
                            if not has_date:
                                next_row_original = df.iloc[idx + 1].values
                                for i, val in enumerate(next_row_original):
                                    if pd.notna(val):
                                        val_str = str(val).strip()
                                        if self.amount_pattern.match(val_str):
                                            clean_amount = self.clean_amount(val_str)
                                            if clean_amount is not None:
                                                amounts.append(clean_amount)
                                                amount_positions.append(i)
                                        elif val_str not in ['NaN', 'nan']:
                                            desc_parts.append(val_str)
                                skip_next_row = True
                        
                        # Clean up description
                        desc = ' '.join(desc_parts)
                        desc = re.sub(r'\s+', ' ', desc).strip()
                        current_transaction['Description'] = desc
                        
                        # Assign amounts based on position and NaN context
                        if amounts:
                            # Last amount is always balance
                            current_transaction['Balance ($)'] = amounts[-1]
                            
                            if len(amounts) > 1:
                                first_amount = amounts[0]
                                first_pos = amount_positions[0]
                                
                                # Check NaN pattern in original row
                                nan_pattern = [pd.isna(val) for val in original_values]
                                
                                # In ANZ statements:
                                # If NaN appears before the amount, it's usually a deposit
                                # If NaN appears after the amount, it's usually a withdrawal
                                nan_before = any(nan_pattern[:first_pos])
                                nan_after = any(nan_pattern[first_pos+1:])
                                
                                if nan_before and not nan_after:
                                    current_transaction['Deposits ($)'] = first_amount
                                elif not nan_before and nan_after:
                                    current_transaction['Withdrawals ($)'] = first_amount
                                else:
                                    # Fallback to CR/DR indicators
                                    if 'CR' in row_text:
                                        current_transaction['Deposits ($)'] = first_amount
                                    else:
                                        current_transaction['Withdrawals ($)'] = first_amount
                    
                    elif current_transaction:
                        # Add content to current transaction
                        desc_parts = []
                        amounts = []
                        
                        for val in original_values:
                            if pd.notna(val):
                                val_str = str(val).strip()
                                if self.amount_pattern.match(val_str):
                                    clean_amount = self.clean_amount(val_str)
                                    if clean_amount is not None:
                                        amounts.append(clean_amount)
                                elif val_str not in ['NaN', 'nan']:
                                    desc_parts.append(val_str)
                        
                        if desc_parts:
                            current_transaction['Description'] += ' ' + ' '.join(desc_parts)
                            current_transaction['Description'] = re.sub(r'\s+', ' ', current_transaction['Description']).strip()
                        
                        if amounts and current_transaction['Balance ($)'] == 0:
                            current_transaction['Balance ($)'] = amounts[-1]
                
                # Don't forget the last transaction
                if current_transaction:
                    all_transactions.append(current_transaction)
            
            # Create final DataFrame
            if all_transactions:
                df = pd.DataFrame(all_transactions)
                
                # Clean up the data
                df = df.dropna(subset=['Date', 'Description'])
                
                # Ensure amounts are float
                for col in ['Withdrawals ($)', 'Deposits ($)', 'Balance ($)']:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
                
                # Sort by date
                try:
                    df['Date'] = pd.to_datetime(df['Date'], format='%d %b')
                    df = df.sort_values('Date')
                except Exception as e:
                    logger.warning(f"Could not sort by date: {str(e)}")
                
                # Reorder columns to standard format
                df = df[self.standard_columns]
                
                return df
            
            return pd.DataFrame(columns=self.standard_columns)
            
        except Exception as e:
            logger.error(f"Error parsing PDF: {str(e)}")
            raise Exception(f"Failed to parse PDF: {str(e)}")
    
    def extract_tables_fallback(self, pdf_path: str) -> pd.DataFrame:
        """
        Alternative method for extracting tables when the primary method fails.
        Uses a different approach to avoid infinite recursion.
        """
        try:
            # Try using lattice mode instead of stream mode
            dfs = tabula.read_pdf(
                pdf_path,
                pages='all',
                lattice=True,  # Use lattice mode instead of stream
                multiple_tables=True,
                guess=False,  # Don't guess table structure
                pandas_options={'header': None}
            )
            
            if not dfs:
                logger.error("No tables found in PDF using fallback method")
                return pd.DataFrame(columns=self.standard_columns)
            
            # Combine all tables into one DataFrame
            combined_df = pd.concat(dfs, ignore_index=True)
            
            # Clean up the data
            combined_df = combined_df.replace('', np.nan).dropna(how='all')
            
            # Try to identify columns based on content
            df_final = pd.DataFrame(columns=self.standard_columns)
            
            for idx, row in combined_df.iterrows():
                transaction = {
                    'Date': '',
                    'Description': '',
                    'Withdrawals ($)': 0.0,
                    'Deposits ($)': 0.0,
                    'Balance ($)': 0.0
                }
                
                # Process each cell in the row
                desc_parts = []
                for val in row:
                    if pd.isna(val):
                        continue
                    
                    val_str = str(val).strip()
                    if self.is_date(val_str):
                        transaction['Date'] = val_str
                    elif self.is_amount(val_str):
                        amount = self.clean_amount(val_str)
                        if amount is not None:
                            if 'CR' in val_str:
                                transaction['Deposits ($)'] = amount
                            elif transaction['Balance ($)'] == 0:
                                transaction['Balance ($)'] = amount
                            else:
                                transaction['Withdrawals ($)'] = amount
                    else:
                        desc_parts.append(val_str)
                
                transaction['Description'] = ' '.join(desc_parts).strip()
                
                # Only add rows that have at least a date or description
                if transaction['Date'] or transaction['Description']:
                    df_final = pd.concat([df_final, pd.DataFrame([transaction])], ignore_index=True)
            
            return df_final
            
        except Exception as e:
            logger.error(f"Error in fallback extraction: {str(e)}")
            return pd.DataFrame(columns=self.standard_columns)

    def extract_tables(self, pdf_path: str) -> pd.DataFrame:
        """
        Main entry point for table extraction.
        Uses parse_pdf_to_excel internally.
        """
        try:
            df = self.parse_pdf_to_excel(pdf_path, fallback_attempted=False)
            
            # Ensure proper data types for amount columns
            for col in ['Withdrawals ($)', 'Deposits ($)', 'Balance ($)']:
                try:
                    # Fixed escape sequence in regex
                    df[col] = pd.to_numeric(df[col].replace(r'[$,]', '', regex=True), errors='coerce').fillna(0.0)
                except Exception as e:
                    logger.warning(f"Error converting column {col} to numeric: {str(e)}")
                    df[col] = 0.0
            
            return df
            
        except Exception as e:
            logger.error(f"Error in table extraction: {str(e)}")
            return pd.DataFrame(columns=self.standard_columns)

    def is_date(self, text: str) -> bool:
        """Check if a string matches ANZ date format."""
        if pd.isna(text):
            return False
        text = str(text).strip().upper()
        return bool(self.date_pattern.match(text))
    
    def is_amount(self, text: str) -> bool:
        """Check if a string matches amount format."""
        if pd.isna(text):
            return False
        text = str(text).strip()
        return bool(self.amount_pattern.match(text))
    
    def clean_amount(self, value: Any) -> Optional[float]:
        """Clean and convert amount strings to float values."""
        if pd.isna(value):
            return None
        try:
            # Remove currency symbols and commas
            cleaned = str(value).replace('$', '').replace(',', '').strip()
            if cleaned:
                return float(cleaned)
        except (ValueError, TypeError):
            pass
        return None 