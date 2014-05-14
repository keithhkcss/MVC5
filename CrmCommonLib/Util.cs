using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Configuration;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Crm.Sdk.Messages;
using System.ServiceModel.Description;
using Microsoft.Crm.Services.Utility;
using Xrm;
using System.Collections.ObjectModel;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;

namespace CrmCommonLib
{
    public class Util
    {
        //******* CRM Service Init (start) *******
        
        #region ***** (1) CRM Service Init *****
        public static OrganizationServiceProxy InitServiceProxy(string mode)
        {
            if (mode == "1")
                return ServiceProxyOnPremise();
            else
                return ServiceProxyOnline();
        }

        public static OrganizationServiceProxy ServiceProxyOnPremise()
        {
            //string OrganizationUrl = "http://devildog.introv.com:4008/WebOrganicDEV/XRMServices/2011/Organization.svc";
            //string OrganizationUrl = ConfigurationSettings.AppSettings["OrganizationUrlOnPremise"];
            //string UserName = ConfigurationSettings.AppSettings["UsernameOnPremise"];
            //string Password = ConfigurationSettings.AppSettings["PasswordOnPremise"];

            string OrganizationUrl = "http://devildog.introv.com:4008/WebOrganicDEV/XRMServices/2011/Organization.svc";
            string UserName = "administrator";
            string Password = "P@ssword321";


            Uri organizationUri = new Uri(OrganizationUrl);

            System.ServiceModel.Description.ClientCredentials credentials = new System.ServiceModel.Description.ClientCredentials();
            credentials.Windows.ClientCredential = new System.Net.NetworkCredential(UserName, Password, "");

            return new OrganizationServiceProxy(organizationUri, null, credentials, null); ;
        }

        public static OrganizationServiceProxy ServiceProxyOnline()
        {
            /*Start For Online Proudction */
            //this is POC only that's why main code is here 

            string OrganizationUrl = "https://WeborganicDev791.crm5.dynamics.com/XRMServices/2011/Organization.svc";
            string UserName = "admin@WeborganicDev791.onmicrosoft.com";
            string Password = "Test1234";

            //string OrganizationUrl = ConfigurationSettings.AppSettings["OrganizationUrlOnline"];
            //string UserName = ConfigurationSettings.AppSettings["UsernameOnline"];
            //string Password = ConfigurationSettings.AppSettings["PasswordOnline"];   

            ClientCredentials credentials = new ClientCredentials(); ;
            credentials.UserName.UserName = UserName;
            credentials.UserName.Password = Password;

            //AuthConfiguration little helper stores all authentication paramters
            AuthConfiguration authconfig = new AuthConfiguration();
            authconfig.Credentials = credentials;

            //DeviceIdManager should have PersistToFile = false;
            ClientCredentials deviceCredentials = DeviceIdManager.LoadOrRegisterDevice("whateveryoulike", "anypassword");
            authconfig.DeviceCredentials = deviceCredentials;

            //set CRM service URL
            authconfig.OrganizationUri = new Uri(OrganizationUrl);

            return new OrganizationServiceProxy(authconfig.OrganizationUri, authconfig.HomeRealmUri, authconfig.Credentials, authconfig.DeviceCredentials);
            /*End For Online Proudction */
        }

        #endregion

        //------------------------------------------------

        #region ***** (2) CRM Service Init (with user login entry) *****
        public static OrganizationServiceProxy InitServiceProxy(string mode, string userID, string password)
        {
            if (mode == "1")
                return ServiceProxyOnPremise(userID, password);
            else
                return ServiceProxyOnline(userID, password);
        }

        public static OrganizationServiceProxy ServiceProxyOnPremise(string userID, string password)
        {
            //string OrganizationUrl = "http://devildog.introv.com:4008/WebOrganicDEV/XRMServices/2011/Organization.svc";
            //string OrganizationUrl = ConfigurationSettings.AppSettings["OrganizationUrlOnPremise"];
            //string UserName = ConfigurationSettings.AppSettings["UsernameOnPremise"];
            //string Password = ConfigurationSettings.AppSettings["PasswordOnPremise"];

            string OrganizationUrl = "http://devildog.introv.com:4008/WebOrganicDEV/XRMServices/2011/Organization.svc";
            string UserName = userID;
            string Password = password;


            Uri organizationUri = new Uri(OrganizationUrl);

            System.ServiceModel.Description.ClientCredentials credentials = new System.ServiceModel.Description.ClientCredentials();
            credentials.Windows.ClientCredential = new System.Net.NetworkCredential(UserName, Password, "");

            return new OrganizationServiceProxy(organizationUri, null, credentials, null); ;
        }

        public static OrganizationServiceProxy ServiceProxyOnline(string userID, string password)
        {
            /*Start For Online Proudction */
            //this is POC only that's why main code is here 

            string OrganizationUrl = "https://WeborganicDev791.crm5.dynamics.com/XRMServices/2011/Organization.svc";
            string UserName = userID;
            string Password = password;

            //string OrganizationUrl = ConfigurationSettings.AppSettings["OrganizationUrlOnline"];
            //string UserName = ConfigurationSettings.AppSettings["UsernameOnline"];
            //string Password = ConfigurationSettings.AppSettings["PasswordOnline"];   

            ClientCredentials credentials = new ClientCredentials(); ;
            credentials.UserName.UserName = UserName;
            credentials.UserName.Password = Password;

            //AuthConfiguration little helper stores all authentication paramters
            AuthConfiguration authconfig = new AuthConfiguration();
            authconfig.Credentials = credentials;

            //DeviceIdManager should have PersistToFile = false;
            ClientCredentials deviceCredentials = DeviceIdManager.LoadOrRegisterDevice("whateveryoulike", "anypassword");
            authconfig.DeviceCredentials = deviceCredentials;


            //set CRM service URL
            authconfig.OrganizationUri = new Uri(OrganizationUrl);

            return new OrganizationServiceProxy(authconfig.OrganizationUri, authconfig.HomeRealmUri, authconfig.Credentials, authconfig.DeviceCredentials);
            /*End For Online Proudction */
        }
        #endregion

        //******* CRM Service Init (end) *******
        


        //******* Data Formating (start) *******
        #region ***** Data Formating *****
        public static string DecimalToText_ZeroToBlank(decimal val, int decPlace)
        {
            if (val == 0)
                return "";
            else
                return Decimal.Round(val, decPlace).ToString();
        }

        public static string OptionValueToText(Entity entity, OptionMetadata[] optionList, string fieldName)
        {
            string returnValue = "";

            if (entity.Attributes.Contains(fieldName))
            {
                OptionSetValue optionValue;

                optionValue = (OptionSetValue)entity.Attributes[fieldName];
                if (optionValue != null)
                    returnValue = Util.GetOptionsetText(optionList, fieldName, optionValue.Value);
                else
                    returnValue = "-";
            }

            if (returnValue != "")
                return returnValue;
            else
                return "-";
        }

        public static string OptionValueToText(OptionSetValue optionValue, Dictionary<int, string> optionList)
        {
            string returnValue = "";

            if (optionValue != null)
                returnValue = optionList[optionValue.Value];
            else
                returnValue = "";

            if (returnValue != "")
                return returnValue;
            else
                return "";
        }

        public static string OptionValueToText(Entity entity, Dictionary<int, string> optionList, string fieldName)
        {
            OptionSetValue optionValue;
            try
            {
                string returnValue = "";

                if (entity.Attributes.Contains(fieldName))
                {
                    optionValue = (OptionSetValue)entity.Attributes[fieldName];
                    if (optionValue != null)
                        returnValue = optionList[optionValue.Value];
                    else
                        returnValue = "";
                }

                if (returnValue != "")
                    return returnValue;
                else
                    return "";
            }
            catch (Exception ex)
            {
                //Console.WriteLine(optionValue.Value.ToString());
                return "-";
            }
        }

        public static string BoolToText(Entity entity, string fieldName)
        {
            string returnValue = "";

            if (entity.Attributes.Contains(fieldName))
            {
                if ((bool)entity.Attributes[fieldName])
                    returnValue = "Yes";
                else
                    returnValue = "No";
            }

            return returnValue;
        }

        public static string DatetimeToText(Entity entity, string fieldName, string formatStr)
        {
            string returnValue = "";

            if (entity.Attributes.Contains(fieldName))
            {
                returnValue = ((DateTime)entity.Attributes[fieldName]).ToString(formatStr);
            }

            if (returnValue != "")
                return returnValue;
            else
                return "";
        }
        #endregion
        //******* Data Formating (end) *******
        
        //******* CRM data retrieve (start) *******
        #region ***** CRM data retrieve *****

        public static ObservableCollection<Entity> GetRecordsFromQuery(IOrganizationService service, QueryExpression query)
        {
            // The number of records per page to retrieve.
            int fetchCount = 5000;
            // Initialize the page number.
            int pageNumber = 1;

            // Set Paging
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = fetchCount;
            query.PageInfo.PageNumber = pageNumber;
            // The current paging cookie. When retrieving the first page, 
            // pagingCookie should be null.
            query.PageInfo.PagingCookie = null;


            // Create the request object.            
            OrganizationRequest request = new OrganizationRequest();
            request.RequestName = "RetrieveMultiple";
            request["Query"] = query;

            ObservableCollection<Entity> entities = new ObservableCollection<Entity>();

            int x = 0;
            while (true)
            {
                // Execute the request.
                OrganizationResponse response = (OrganizationResponse)service.Execute(request);
                EntityCollection results = (EntityCollection)response["EntityCollection"];

                if (results != null)
                {
                    if (results.Entities != null)
                    {
                        // Retrieve all records from the result set.
                        foreach (Entity e in results.Entities)
                        {
                            x++;
                            entities.Add(e);
                        }
                    }

                    // Check for more records, if it returns true.
                    if (results.MoreRecords)
                    {
                        //resultsContact
                        // Increment the page number to retrieve the next page.
                        query.PageInfo.PageNumber++;
                        // Set the paging cookie to the paging cookie returned from current results.
                        query.PageInfo.PagingCookie = results.PagingCookie;
                    }
                    else
                    {
                        // If no more records are in the result nodes, exit the loop.
                        break;
                    }
                }

            }

            if (entities.Count > 0)
                return entities;
            else
                return null;
        }

        public static Dictionary<int, string> LoadStateSetValue(IOrganizationService service, string entityName, string attributeName)
        {
            RetrieveAttributeRequest request = new RetrieveAttributeRequest();
            request.RequestName = "RetrieveAttribute";
            request["EntityLogicalName"] = entityName;
            request["LogicalName"] = attributeName;
            request["MetadataId"] = Guid.Empty;
            request["RetrieveAsIfPublished"] = true;

            RetrieveAttributeResponse response = (RetrieveAttributeResponse)service.Execute(request);
            StateAttributeMetadata picklist = (StateAttributeMetadata)response.AttributeMetadata;

            Dictionary<int, string> optionDict = new Dictionary<int, string>();

            foreach (OptionMetadata option in picklist.OptionSet.Options)
            {
                optionDict.Add((int)option.Value, option.Label.UserLocalizedLabel.Label);
            }

            return optionDict;
        }

        public static Dictionary<int, string> LoadOptionsetValue(IOrganizationService service, string entityName, string attributeName)
        {
            RetrieveAttributeRequest request = new RetrieveAttributeRequest();
            request.RequestName = "RetrieveAttribute";
            request["EntityLogicalName"] = entityName;
            request["LogicalName"] = attributeName;
            request["MetadataId"] = Guid.Empty;
            request["RetrieveAsIfPublished"] = true;

            RetrieveAttributeResponse response = (RetrieveAttributeResponse)service.Execute(request);
            PicklistAttributeMetadata picklist = (PicklistAttributeMetadata)response.AttributeMetadata;

            Dictionary<int, string> optionDict = new Dictionary<int, string>();

            foreach (OptionMetadata option in picklist.OptionSet.Options)
            {
                optionDict.Add((int)option.Value, option.Label.UserLocalizedLabel.Label);
            }

            return optionDict;
        }

        public static Dictionary<int, string> LoadOptionsetValue(IOrganizationService service, string optionsetName)
        {
            //private Dictionary<int, string> typeOptionSet = new Dictionary<int, string>();
            Dictionary<int, string> optionDict = new Dictionary<int, string>();

            try
            {
                Microsoft.Xrm.Sdk.Messages.RetrieveOptionSetRequest retrieveOptionSetRequest =
                    new Microsoft.Xrm.Sdk.Messages.RetrieveOptionSetRequest
                    {
                        Name = optionsetName
                    };

                // Execute the request.
                Microsoft.Xrm.Sdk.Messages.RetrieveOptionSetResponse retrieveOptionSetResponse = (Microsoft.Xrm.Sdk.Messages.RetrieveOptionSetResponse)service.Execute(retrieveOptionSetRequest);

                // Access the retrieved OptionSetMetadata.
                OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;

                // Get the current options list for the retrieved attribute.
                OptionMetadata[] optionList = retrievedOptionSetMetadata.Options.ToArray();
                foreach (OptionMetadata optionMetadata in optionList)
                {
                    optionDict.Add((int)optionMetadata.Value, optionMetadata.Label.UserLocalizedLabel.Label);
                }
            }
            catch (Exception)
            {
                throw;
            }

            return optionDict;
        }

        public static string GetOptionsetText(Entity entity, IOrganizationService service, string optionsetName, int optionsetValue)
        {
            string optionsetSelectedText = string.Empty;
            try
            {
                Microsoft.Xrm.Sdk.Messages.RetrieveOptionSetRequest retrieveOptionSetRequest =
                    new Microsoft.Xrm.Sdk.Messages.RetrieveOptionSetRequest
                    {
                        Name = optionsetName
                    };

                // Execute the request.
                Microsoft.Xrm.Sdk.Messages.RetrieveOptionSetResponse retrieveOptionSetResponse = (Microsoft.Xrm.Sdk.Messages.RetrieveOptionSetResponse)service.Execute(retrieveOptionSetRequest);

                // Access the retrieved OptionSetMetadata.
                OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;

                // Get the current options list for the retrieved attribute.
                OptionMetadata[] optionList = retrievedOptionSetMetadata.Options.ToArray();
                foreach (OptionMetadata optionMetadata in optionList)
                {
                    if (optionMetadata.Value == optionsetValue)
                    {
                        optionsetSelectedText = optionMetadata.Label.UserLocalizedLabel.Label.ToString();
                        break;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            return optionsetSelectedText;
        }

        public static string GetOptionsetText(OptionMetadata[] optionList, string optionsetName, int optionsetValue)
        {
            string optionsetSelectedText = string.Empty;
            if (optionList != null)
            {
                foreach (OptionMetadata optionMetadata in optionList)
                {
                    if (optionMetadata.Value == optionsetValue)
                    {
                        optionsetSelectedText = optionMetadata.Label.UserLocalizedLabel.Label.ToString();
                        break;
                    }
                }
            }
            return optionsetSelectedText;
        }

        public static OptionMetadata[] GetOptionsetMetadata(IOrganizationService service, string optionsetName)
        {
            OptionMetadata[] optionList = null;

            try
            {
                Microsoft.Xrm.Sdk.Messages.RetrieveOptionSetRequest retrieveOptionSetRequest =
                    new Microsoft.Xrm.Sdk.Messages.RetrieveOptionSetRequest
                    {
                        Name = optionsetName
                    };

                // Execute the request.
                Microsoft.Xrm.Sdk.Messages.RetrieveOptionSetResponse retrieveOptionSetResponse = (Microsoft.Xrm.Sdk.Messages.RetrieveOptionSetResponse)service.Execute(retrieveOptionSetRequest);

                // Access the retrieved OptionSetMetadata.
                OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;

                // Get the current options list for the retrieved attribute.
                optionList = retrievedOptionSetMetadata.Options.ToArray();
            }
            catch (Exception)
            {
                throw;
            }
            return optionList;
        }

        public static Entity GetSingleEntity(IOrganizationService service, string EntityName, string[] ConditionAttrNames, object[] ConditionAttrValues, string[] Cols)
        {
            ConditionExpression conditions;
            FilterExpression filters = new FilterExpression();

            //ConditionAttrNames.Length

            for (int i = 0; i < ConditionAttrNames.Length; i++)
            {
                conditions = new ConditionExpression();
                conditions.AttributeName = ConditionAttrNames[i];
                conditions.Operator = ConditionOperator.Equal;
                conditions.Values.Add(ConditionAttrValues[i]);

                filters.Conditions.Add(conditions);
            }

            ColumnSet cols = new ColumnSet(Cols);
            //OrderExpression orders = new OrderExpression("new_name", OrderType.Ascending);            

            QueryExpression query = new QueryExpression();
            query.EntityName = EntityName;
            query.Criteria.AddFilter(filters);
            query.ColumnSet = cols;
            //query.Orders.Add(orders);

            // Create the request object.            
            OrganizationRequest request = new OrganizationRequest();
            request.RequestName = "RetrieveMultiple";
            request["Query"] = query;

            // Execute the request.
            OrganizationResponse response = (OrganizationResponse)service.Execute(request);
            EntityCollection results = (EntityCollection)response["EntityCollection"];

            if (results.Entities.Count > 0)
            {
                return results.Entities[0];
            }
            else
            {
                return null;
            }
        }

        public static ObservableCollection<Entity> GetMultipleEntity(IOrganizationService service, string EntityName, string[] ConditionAttrNames, object[] ConditionAttrValues,
            object[] ConditionOperatorVals, LogicalOperator logicalOperator, string[] Cols)
        {
            // The number of records per page to retrieve.
            int fetchCount = 5000;
            // Initialize the page number.
            int pageNumber = 1;

            ConditionExpression conditions;
            FilterExpression filters = new FilterExpression();

            if (ConditionAttrNames != null)
            {
                for (int i = 0; i < ConditionAttrNames.Length; i++)
                {
                    conditions = new ConditionExpression();
                    conditions.AttributeName = ConditionAttrNames[i];
                    if (ConditionOperatorVals == null)
                    {
                        conditions.Operator = ConditionOperator.Equal;
                    }
                    else
                    {
                        conditions.Operator = (ConditionOperator)ConditionOperatorVals[i];
                    }

                    if ((ConditionOperator)ConditionOperatorVals[i] != ConditionOperator.Null && (ConditionOperator)ConditionOperatorVals[i] != ConditionOperator.NotNull)
                    {
                        conditions.Values.Add(ConditionAttrValues[i]);
                    }

                    filters.Conditions.Add(conditions);
                }
            }

            if (logicalOperator != null)
                filters.FilterOperator = logicalOperator;

            ColumnSet cols = new ColumnSet(Cols);
            //OrderExpression orders = new OrderExpression("new_name", OrderType.Ascending);            

            QueryExpression query = new QueryExpression();
            query.EntityName = EntityName;
            if (ConditionAttrNames != null)
            {
                query.Criteria.AddFilter(filters);
            }
            query.ColumnSet = cols;
            //query.Orders.Add(orders);

            // Set Paging
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = fetchCount;
            query.PageInfo.PageNumber = pageNumber;
            // The current paging cookie. When retrieving the first page, 
            // pagingCookie should be null.
            query.PageInfo.PagingCookie = null;

            // Create the request object.            
            OrganizationRequest request = new OrganizationRequest();
            request.RequestName = "RetrieveMultiple";
            request["Query"] = query;

            ObservableCollection<Entity> entities = new ObservableCollection<Entity>();

            int x = 0;
            while (true)
            {
                // Execute the request.
                OrganizationResponse response = (OrganizationResponse)service.Execute(request);
                EntityCollection results = (EntityCollection)response["EntityCollection"];

                if (results != null)
                {
                    if (results.Entities != null)
                    {
                        // Retrieve all records from the result set.
                        foreach (Entity e in results.Entities)
                        {
                            x++;
                            entities.Add(e);
                        }
                    }

                    // Check for more records, if it returns true.
                    if (results.MoreRecords)
                    {
                        //resultsContact
                        // Increment the page number to retrieve the next page.
                        query.PageInfo.PageNumber++;
                        // Set the paging cookie to the paging cookie returned from current results.
                        query.PageInfo.PagingCookie = results.PagingCookie;
                    }
                    else
                    {
                        // If no more records are in the result nodes, exit the loop.
                        break;
                    }
                }

            }

            if (entities.Count > 0)
            {
                return entities;
            }
            else
            {
                return null;
            }
        }
        #endregion
        //******* CRM data retrieve (end) *******

        //******* Log Writer (start) *******
        #region ***** Log Writer *****
        public static TextWriter CreateLogWriter(string prefix)
        {
            string folderPath = Directory.GetCurrentDirectory() + "/Log";
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            string timeString = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            StreamWriter writer;
            writer = File.CreateText(folderPath + "/" + prefix + timeString + ".txt");

            return writer;
        }

        public static void WriteLog(TextWriter writer, string log)
        {
            string timeString = DateTime.Now.ToString("HH:mm:ss");

            writer.WriteLine(timeString + "\t" + log);
        }
        #endregion
        //******* Log Writer (end) *******

        //******* CSV Generation (start) *******
        #region ***** CSV Generation *****

        public static StreamWriter CreateCSVWriter(string fileName)
        {
            string folderPath = Directory.GetCurrentDirectory() + "/CSV";
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            //string timeString = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            StreamWriter writer;
            writer = File.CreateText(folderPath + "/" + fileName + ".csv");

            return writer;
        }

        public static string CreateFolder(string FolderName)
        {
            if (!Directory.Exists(FolderName))
                Directory.CreateDirectory(FolderName);

            return FolderName;
        }


        public string ConvertToCSV(CsvRow csvRow)
        {
            StringBuilder builder = new StringBuilder();
            bool firstColumn = true;
            foreach (string value in csvRow)
            {
                // Add separator if this isn't the first value
                if (!firstColumn)
                    builder.Append(',');
                // Implement special handling for values that contain comma or quote
                // Enclose in quotes and double up any double quotes
                if (value.IndexOfAny(new char[] { '"', ',' }) != -1)
                    builder.AppendFormat("\"{0}\"", value.Replace("\"", "\"\""));
                else
                    builder.Append(value);
                firstColumn = false;
            }
            return builder.ToString();
        }

        /// <summary>
        /// Class to store one CSV row
        /// </summary>
        public class CsvRow : List<string>
        {
            public string LineText { get; set; }
        }

        /// <summary>
        /// Class to write data to a CSV file
        /// </summary>
        public class CsvFileWriter : StreamWriter
        {
            public CsvFileWriter(Stream stream)
                : base(stream)
            {
            }

            public CsvFileWriter(string filename)
                : base(filename)
            {
            }

            public CsvFileWriter(string filename, bool append, Encoding encoding)
                : base(filename, append, encoding)
            {
            }

            /// <summary>
            /// Writes a single row to a CSV file.
            /// </summary>
            /// <param name="row">The row to be written</param>
            public void WriteRow(CsvRow row)
            {
                StringBuilder builder = new StringBuilder();
                bool firstColumn = true;
                foreach (string value in row)
                {
                    // Add separator if this isn't the first value
                    if (!firstColumn)
                        builder.Append(',');
                    // Implement special handling for values that contain comma or quote
                    // Enclose in quotes and double up any double quotes
                    int n;
                    bool isNumeric = int.TryParse(value, out n);

                    DateTime d;
                    bool isDate = DateTime.TryParse(value, out d);

                    //if(value.Length > 7)
                    //    isDate = DateTime.TryParse(value, out d);

                    if (!isNumeric && !isDate)
                    {
                        //if (value.IndexOfAny(new char[] { '"', ',' }) != -1)
                        //    builder.AppendFormat("\"{0}\"", value.Replace("\"", "\"\""));
                        //else
                        //    builder.Append(value);

                        if(value != null)
                            builder.AppendFormat("\"{0}\"", value.Replace("\"", "\"\""));
                        else
                            builder.Append("");
                    }
                    else
                    {
                        builder.Append(value);
                    }

                    firstColumn = false;
                }
                row.LineText = builder.ToString();
                WriteLine(row.LineText);
            }
        }
        #endregion
        //******* CSV Generation (end) *******

        //******* Common functions (start) *******
        public static int MonthDiff(DateTime startDate, DateTime endDate)
        {
            int monthsApart = 12 * (startDate.Year - endDate.Year) + startDate.Month - endDate.Month;
            return Math.Abs(monthsApart);
        }

        public static DateTime SetFromDate(DateTime d)
        {
            DateTime tmpDate = new DateTime(d.Year, d.Month, 1);
            return tmpDate;
        }

        public static DateTime SetToDate(DateTime d)
        {
            int iYear = 0;
            int iMonth = 0;
            int iDay = 1;

            if (d.Month == 12)
            {
                iYear = d.Year + 1;
                iMonth = 1;
            }
            else
            {
                iYear = d.Year;
                iMonth = d.Month + 1;
            }

            DateTime tmpDate = new DateTime(iYear, iMonth, iDay);
            DateTime toDate = tmpDate.AddDays(-1);

            return toDate;
        }
        //******* Common functions (end) *******
    }
}
