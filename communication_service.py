"""
Communication Service Module for sending text and email messages via myKaarma API.

This module handles:
1. Template loading and processing
2. Text message sending
3. Email message sending
4. Error handling and logging
"""

import os
import requests
import json
import logging
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import Dict, Optional, Tuple
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

class CommunicationService:
    """Service class for handling text and email communications via myKaarma API."""
    
    def __init__(self):
        self.base_url = os.getenv('MYKAARMA_BASE_URL')
        self.username = os.getenv('MYKAARMA_USERNAME')
        self.password = os.getenv('MYKAARMA_PASSWORD')
        self.auth = (self.username, self.password)
        
        # Load templates
        self.email_template = self._load_template('templates/email_template.txt')
        self.text_template = self._load_template('templates/text_template.txt')
        
        # Cache for default user UUIDs to avoid repeated API calls
        self._default_user_cache = {}
        
        # Setup logging
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)
    
    def _load_template(self, template_path: str) -> str:
        """Load template content from file."""
        try:
            with open(template_path, 'r', encoding='utf-8') as file:
                return file.read()
        except FileNotFoundError:
            self.logger.error(f"Template file not found: {template_path}")
            return ""
        except Exception as e:
            self.logger.error(f"Error loading template {template_path}: {e}")
            return ""
    
    def _format_date_time(self, date_str: str, time_str: str, date_formats: Dict[str, str]) -> Dict[str, str]:
        """
        Format date and time according to specified formats.
        
        Args:
            date_str: Date string in YYYY-MM-DD format
            time_str: Time string in HH:MM:SS format
            date_formats: Dictionary mapping variable names to format patterns
            
        Returns:
            Dictionary of formatted date/time values
        """
        try:
            # Combine date and time for parsing
            datetime_str = f"{date_str} {time_str}"
            dt = datetime.strptime(datetime_str, '%Y-%m-%d %H:%M:%S')
            
            formatted_values = {}
            
            for var_name, format_pattern in date_formats.items():
                if var_name == '_appt_date':
                    # Format: EEEE, MMMM dd, yyyy -> Monday, January 15, 2024
                    if format_pattern == 'EEEE, MMMM dd, yyyy':
                        formatted_values[var_name] = dt.strftime('%A, %B %d, %Y')
                    else:
                        formatted_values[var_name] = dt.strftime('%Y-%m-%d')  # fallback
                elif var_name == '_appt_start_time':
                    # Format: hh:mm a -> 09:30 AM
                    if format_pattern == 'hh:mm a':
                        formatted_values[var_name] = dt.strftime('%I:%M %p')
                    else:
                        formatted_values[var_name] = dt.strftime('%H:%M:%S')  # fallback
                        
            return formatted_values
            
        except Exception as e:
            self.logger.error(f"Error formatting date/time: {e}")
            # Return original values as fallback
            return {
                '_appt_date': date_str,
                '_appt_start_time': time_str
            }
    
    def _parse_email_template(self, template: str) -> Tuple[str, str, Dict[str, str]]:
        """
        Parse the XML-like email template format.
        
        Args:
            template: The email template string in XML format
            
        Returns:
            Tuple of (subject, body, date_formats)
        """
        try:
            # Parse the XML
            root = ET.fromstring(template.strip())
            
            subject = root.find('subject')
            body = root.find('body')
            
            subject_text = subject.text if subject is not None else ""
            body_text = body.text if body is not None else ""
            
            # Extract date format specifications
            date_formats = {}
            for date_format_elem in root.findall('date_format'):
                if date_format_elem.text:
                    parts = date_format_elem.text.split('#')
                    if len(parts) == 2:
                        format_pattern, var_name = parts
                        date_formats[var_name] = format_pattern
            
            return subject_text, body_text, date_formats
            
        except Exception as e:
            self.logger.error(f"Error parsing email template: {e}")
            # Fallback to simple parsing
            return template, "", {}
    
    def _parse_text_template(self, template: str) -> Tuple[str, Dict[str, str]]:
        """
        Parse the XML-like text template format.
        
        Args:
            template: The text template string in XML format
            
        Returns:
            Tuple of (body, date_formats)
        """
        try:
            # Parse the XML
            root = ET.fromstring(template.strip())
            
            body = root.find('body')
            body_text = body.text if body is not None else ""
            
            # Extract date format specifications
            date_formats = {}
            for date_format_elem in root.findall('date_format'):
                if date_format_elem.text:
                    parts = date_format_elem.text.split('#')
                    if len(parts) == 2:
                        format_pattern, var_name = parts
                        date_formats[var_name] = format_pattern
            
            return body_text, date_formats
            
        except Exception as e:
            self.logger.error(f"Error parsing text template: {e}")
            # Fallback to simple parsing
            return template, {}
    
    def _process_template(self, template: str, variables: Dict[str, str]) -> str:
        """
        Process template by replacing variables with underscore prefix format.
        
        Args:
            template: The template string
            variables: Dictionary of variable name -> value mappings
            
        Returns:
            Processed template string
        """
        processed = template
        for var_name, value in variables.items():
            # Variables now use underscore prefix format
            processed = processed.replace(var_name, str(value))
        
        return processed
    
    def _extract_email_subject_and_body(self, email_template: str, variables: Dict[str, str]) -> Tuple[str, str]:
        """
        Extract and process subject and body from email template.
        
        Args:
            email_template: Full email template in XML format
            variables: Dictionary of variable name -> value mappings
            
        Returns:
            Tuple of (processed_subject, processed_body)
        """
        try:
            # Parse the XML template
            subject_template, body_template, date_formats = self._parse_email_template(email_template)
            
            # Process templates with variables
            processed_subject = self._process_template(subject_template, variables)
            processed_body = self._process_template(body_template, variables)
            
            return processed_subject, processed_body
            
        except Exception as e:
            self.logger.error(f"Error extracting email subject and body: {e}")
            # Fallback to processing the entire template as body
            processed = self._process_template(email_template, variables)
            return "Appointment Confirmation", processed
    
    def _extract_text_body(self, text_template: str, variables: Dict[str, str]) -> str:
        """
        Extract and process body from text template.
        
        Args:
            text_template: Full text template in XML format
            variables: Dictionary of variable name -> value mappings
            
        Returns:
            Processed text body
        """
        try:
            # Parse the XML template
            body_template, date_formats = self._parse_text_template(text_template)
            
            # Process template with variables
            processed_body = self._process_template(body_template, variables)
            
            return processed_body
            
        except Exception as e:
            self.logger.error(f"Error extracting text body: {e}")
            # Fallback to processing the entire template as body
            processed = self._process_template(text_template, variables)
            return processed
    
    def get_default_dealer_associate(self, department_uuid: str) -> Optional[str]:
        """
        Get the default dealer associate user UUID for a department from myKaarma API.
        
        Args:
            department_uuid: The department UUID
            
        Returns:
            Default user UUID if found, None otherwise
        """
        # Check cache first
        if department_uuid in self._default_user_cache:
            return self._default_user_cache[department_uuid]
        
        try:
            url = f"{self.base_url}/manage/v2/department/{department_uuid}/dealerAssociate/default"
            headers = {
                "accept": "application/json"
            }
            
            response = requests.get(url, auth=self.auth, headers=headers, timeout=30)
            response.raise_for_status()
            
            data = response.json()
            
            # Check for errors in response
            errors = data.get('errors', [])
            if errors:
                error_messages = [error.get('errorMessage', 'Unknown error') for error in errors]
                self.logger.error(f"API returned errors for department {department_uuid}: {error_messages}")
                return None
            
            # Extract userUuid from dealerAssociate
            dealer_associate = data.get('dealerAssociate', {})
            user_uuid = dealer_associate.get('userUuid')
            
            if user_uuid:
                # Cache the result
                self._default_user_cache[department_uuid] = user_uuid
                return user_uuid
            else:
                self.logger.warning(f"No default dealer associate userUuid found for department {department_uuid}")
                return None
             
        except requests.exceptions.RequestException as e:
            self.logger.error(f"API request failed while fetching default dealer associate: {e}")
            return None
        except Exception as e:
            self.logger.error(f"Unexpected error fetching default dealer associate for department {department_uuid}: {e}")
            return None
    
    def send_text_message(self, 
                         department_uuid: str,
                         user_uuid: str,
                         customer_uuid: str,
                         template_variables: Dict[str, str],
                         appt_date: str = None,
                         appt_time: str = None,
                         add_tcpa_footer: bool = True,
                         add_signature: bool = True,
                         add_footer: bool = True) -> Dict:
        """
        Send a text message to a customer using myKaarma API.
        API will automatically use customer's preferred phone number from myKaarma records.
        
        Args:
            department_uuid: Dealer department UUID
            user_uuid: User UUID (dealer associate)
            customer_uuid: Customer UUID
            template_variables: Variables for template processing
            add_tcpa_footer: Whether to add TCPA footer
            add_signature: Whether to add signature
            add_footer: Whether to add footer
            
        Returns:
            API response dictionary
        """
        try:
            # Add date/time formatting if provided
            if appt_date and appt_time:
                formatted_datetime = self._format_date_time(appt_date, appt_time, {
                    '_appt_date': 'EEEE, MMMM dd, yyyy',
                    '_appt_start_time': 'hh:mm a'
                })
                template_variables.update(formatted_datetime)
            
            # Process template and extract body
            message_body = self._extract_text_body(self.text_template, template_variables)
            
            # Prepare API request
            url = f"{self.base_url}/communications/department/{department_uuid}/user/{user_uuid}/customer/{customer_uuid}/message"
            
            headers = {
                "accept": "application/json",
                "Content-Type": "application/json"
            }
            
            payload = {
                "messageAttributes": {
                    "body": message_body,
                    "isManual": False,  # Automated message
                    "protocol": "TEXT",  
                    "type": "OUTGOING",
                    "messageType": "S",  # "S" for Sent
                    "isRead": False,
                    "messagePurpose": "AC"  # "AC" for Appointment Confirmation
                },
                "messageSendingAttributes": {
                    "sendSynchronously": True,  # Send synchronously for immediate feedback
                    "addTCPAFooter": add_tcpa_footer,
                    "addSignature": add_signature,
                    "addFooter": add_footer,
                    "sendVCard": False
                }
            }
            
            response = requests.post(url, json=payload, auth=self.auth, headers=headers)
            response.raise_for_status()
            
            result = response.json()
            
            return result
            
        except requests.exceptions.RequestException as e:
            self.logger.error(f"API request failed while sending text: {e}")
            return {"status": "FAILED", "error": str(e)}
        except Exception as e:
            self.logger.error(f"Unexpected error while sending text: {e}")
            return {"status": "FAILED", "error": str(e)}
    
    def send_email_message(self,
                          department_uuid: str,
                          user_uuid: str,
                          customer_uuid: str,
                          template_variables: Dict[str, str],
                          appt_date: str = None,
                          appt_time: str = None,
                          add_signature: bool = True,
                          add_footer: bool = True) -> Dict:
        """
        Send an email message to a customer using myKaarma API.
        API will automatically use customer's preferred email address from myKaarma records.
        
        Args:
            department_uuid: Dealer department UUID
            user_uuid: User UUID (dealer associate)
            customer_uuid: Customer UUID
            template_variables: Variables for template processing
            add_signature: Whether to add signature
            add_footer: Whether to add footer
            
        Returns:
            API response dictionary
        """
        try:
            # Add date/time formatting if provided
            if appt_date and appt_time:
                formatted_datetime = self._format_date_time(appt_date, appt_time, {
                    '_appt_date': 'EEEE, MMMM dd, yyyy',
                    '_appt_start_time': 'hh:mm a'
                })
                template_variables.update(formatted_datetime)
            
            # Process template and extract subject/body
            subject, body = self._extract_email_subject_and_body(self.email_template, template_variables)
            
            # Remove newlines from email body as done in AppointmentCommunicationService.java
            body = body.replace('\n', '')
            
            # Prepare API request
            url = f"{self.base_url}/communications/department/{department_uuid}/user/{user_uuid}/customer/{customer_uuid}/message"
            
            headers = {
                "accept": "application/json",
                "Content-Type": "application/json"
            }
            
            payload = {
                "messageAttributes": {
                    "body": body,
                    "subject": subject,
                    "isManual": False,  # Automated message
                    "protocol": "EMAIL",  # Use "E" for email messages, not "EMAIL"
                    "type": "OUTGOING",
                    "messageType": "S",  # "S" for Sent
                    "isRead": False,
                    "messagePurpose": "AC"  # "AC" for Appointment Confirmation
                },
                "messageSendingAttributes": {
                    "sendSynchronously": True,  # Send synchronously for immediate feedback
                    "addSignature": add_signature,
                    "addFooter": add_footer
                }
            }
            
            
            response = requests.post(url, json=payload, auth=self.auth, headers=headers)
            response.raise_for_status()
            
            result = response.json()
            
            return result
            
        except requests.exceptions.RequestException as e:
            self.logger.error(f"API request failed while sending email: {e}")
            return {"status": "FAILED", "error": str(e)}
        except Exception as e:
            self.logger.error(f"Unexpected error while sending email: {e}")
            return {"status": "FAILED", "error": str(e)}
    
    def send_appointment_notifications(self,
                                     department_uuid: str,
                                     customer_uuid: str,
                                     customer_firstname: str,
                                     customer_lastname: str,
                                     dealer_name: str,
                                     appt_date: str = None,
                                     appt_time: str = None,
                                     user_uuid: Optional[str] = None,
                                     send_text: bool = True,
                                     send_email: bool = True) -> Dict:
        """
        Send appointment notification via both text and email.
        API will automatically use customer's preferred contact information from myKaarma records.
        
        Args:
            department_uuid: Dealer department UUID
            customer_uuid: Customer UUID
            customer_name: Customer's name
            dealership_name: Dealership name
            dealership_phone: Dealership phone number
            user_uuid: User UUID (dealer associate) - if None, will fetch default dealer associate
            send_text: Whether to send text notification
            send_email: Whether to send email notification
            
        Returns:
            Dictionary with results for both text and email
        """
        template_variables = {
            "_customer_firstname": customer_firstname,
            "_customer_lastname": customer_lastname,
            "_dealer_name": dealer_name,
            "_appt_date": appt_date or "",
            "_appt_start_time": appt_time or ""
        }
        
        # If no user_uuid provided, fetch the default dealer associate
        if not user_uuid:
            user_uuid = self.get_default_dealer_associate(department_uuid)
            if not user_uuid:
                return {
                    "text_result": {"status": "FAILED", "reason": "Could not fetch default dealer associate"},
                    "email_result": {"status": "FAILED", "reason": "Could not fetch default dealer associate"},
                    "overall_status": "FAILED"
                }
        
        results = {
            "text_result": None,
            "email_result": None,
            "overall_status": "SUCCESS"
        }
        
        # Send text message if requested
        if send_text:
            text_result = self.send_text_message(
                department_uuid=department_uuid,
                user_uuid=user_uuid,
                customer_uuid=customer_uuid,
                template_variables=template_variables,
                appt_date=appt_date,
                appt_time=appt_time
            )
            results["text_result"] = text_result
            if text_result.get("status") == "FAILED":
                results["overall_status"] = "PARTIAL_FAILED"
        
        # Send email if requested
        if send_email:
            email_result = self.send_email_message(
                department_uuid=department_uuid,
                user_uuid=user_uuid,
                customer_uuid=customer_uuid,
                template_variables=template_variables,
                appt_date=appt_date,
                appt_time=appt_time
            )
            results["email_result"] = email_result
            if email_result.get("status") == "FAILED":
                if results["overall_status"] == "PARTIAL_FAILED":
                    results["overall_status"] = "FAILED"
                else:
                    results["overall_status"] = "PARTIAL_FAILED"
        
        return results


# Convenience function for easy import
def create_communication_service() -> CommunicationService:
    """Create and return a CommunicationService instance."""
    return CommunicationService()
