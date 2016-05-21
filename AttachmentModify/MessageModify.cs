// AttachmentModify
// ----------------------------------------------------------
// Example for intercepting email messages in an Exchange 2010 transport queue
// 
// The example intercepts messages sent from a configurable email address(es)
// and checks the mail message for attachments have filename in to format
// 
//      WORKBOOK_{GUID}
//
// Changing the filename of the attachments makes it easier for the information worker
// to identify the reports in the emails and in the file system as well.
//
// ----------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;

// the lovely Exchange 
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Smtp;
using Microsoft.Exchange.Data.Transport.Email;
using Microsoft.Exchange.Data.Transport.Routing;

namespace SFTools.Messaging.AttachmentModify
{
    #region Message Modifier Factory

    /// <summary>
    /// Message Modifier Factory
    /// </summary>
    public class MessageModifierFactory : RoutingAgentFactory
    {
        /// <summary>
        /// Instance of our transport agent configuration
        /// This is for a later implementation
        /// </summary>
        private MessageModifierConfig messageModifierConfig = new MessageModifierConfig();

        /// <summary>
        /// Returns an instance of the agent
        /// </summary>
        /// <param name="server">The SMTP Server</param>
        /// <returns>The Transport Agent</returns>
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new MessageModifier(messageModifierConfig);
        }
    }

    #endregion

    #region Message Modifier Routing Agent
    
    /// <summary>
    /// The Message Modifier Routing Agent for modifying an email message
    /// </summary>
    public class MessageModifier : RoutingAgent
    {
        // The agent uses the fileLock object to synchronize access to the log file
        private object fileLock = new object();

        /// <summary>
        /// The current MailItem the transport agent is handling
        /// </summary>
        private MailItem mailItem;

        /// <summary>
        /// This context to allow Exchange to continue processing a message
        /// </summary>
        private AgentAsyncContext agentAsyncContext;

        /// <summary>
        /// Transport agent configuration
        /// </summary>
        private MessageModifierConfig messageModifierConfig;

        /// <summary>
        /// Constructor for the MessageModifier class
        /// </summary>
        /// <param name="messageModifierConfig">Transport Agent configuration</param>
        public MessageModifier(MessageModifierConfig messageModifierConfig)
        {
            // Set configuration
            this.messageModifierConfig = messageModifierConfig;

            // Register an OnRoutedMessage event handler
            this.OnRoutedMessage += OnRoutedMessageHandler;
        }

        /// <summary>
        /// Event handler for OnRoutedMessage event
        /// </summary>
        /// <param name="source">Routed Message Event Source</param>
        /// <param name="args">Queued Message Event Arguments</param>
        void OnRoutedMessageHandler(RoutedMessageEventSource source, QueuedMessageEventArgs args)
        {
            lock (fileLock)
            {
                try
                {

                    this.mailItem = args.MailItem;
                    this.agentAsyncContext = this.GetAgentAsyncContext();

                    // Get the folder for accessing the config file
                    string dllDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                    // Fetch the from address from the current mail item
                    RoutingAddress fromAddress = this.mailItem.FromAddress;

                    Boolean boWorkbookFound = false;    // We just want to modifiy subjects when we modified an attachement first

                    #region External Receive Connector Example

                    // CHeck first, if the mail item does have a ReceiveConnectorName property first to prevent ugly things to happen
                    if (mailItem.Properties.ContainsKey("Microsoft.Exchange.Transport.ReceiveConnectorName"))
                    {
                        // This is just an example, if you want to do something with a mail item which has been received via a named external receive connector
                        if (mailItem.Properties["Microsoft.Exchange.Transport.ReceiveConnectorName"].ToString().ToLower() == "externalreceiveconnectorname")
                        {
                            // do something fancy with the email
                        }
                    }

                    #endregion

                    RoutingAddress catchAddress;

                    // Check, if we have any email addresses configured to look for
                    if (this.messageModifierConfig.AddressMap.Count > 0)
                    {
                        // Now lets check, if the sender address can be found in the dictionary
                        if (this.messageModifierConfig.AddressMap.TryGetValue(fromAddress.ToString().ToLower(), out catchAddress))
                        {
                            // Sender address found, now check if we have attachments to handle
                            if (this.mailItem.Message.Attachments.Count != 0)
                            {
                                // Get all attachments
                                AttachmentCollection attachments = this.mailItem.Message.Attachments;

                                // Modify each attachment
                                for (int count = 0; count < this.mailItem.Message.Attachments.Count; count++)
                                {
                                    // Get attachment
                                    Attachment attachment = this.mailItem.Message.Attachments[count];

                                    // We will only transform attachments which start with "WORKBOOK_"
                                    if (attachment.FileName.StartsWith("WORKBOOK_"))
                                    {
                                        // Create a new filename for the attachment
                                        // [MODIFIED SUBJECT]-[NUMBER].[FILEEXTENSION]
                                        String newFileName = MakeValidFileName(string.Format("{0}-{1}{2}", ModifiySubject(this.mailItem.Message.Subject.Trim()), count + 1, Path.GetExtension(attachment.FileName)));

                                        // Change the filename of the attachment
                                        this.mailItem.Message.Attachments[count].FileName = newFileName;

                                        // Yes we have changed the attachment. Therefore we want to change the subject as well.
                                        boWorkbookFound = true;
                                    }
                                }

                                // Have changed any attachments?
                                if (boWorkbookFound)
                                {
                                    // Then let's change the subject as well
                                    this.mailItem.Message.Subject = ModifiySubject(this.mailItem.Message.Subject);
                                }
                            }
                        }
                    }
                }
                catch (System.IO.IOException ex)
                {
                    // oops
                    Debug.WriteLine(ex.ToString());
                    this.agentAsyncContext.Complete();
                }
                finally
                {
                    // We are done
                    this.agentAsyncContext.Complete();
                }
            }

            // Return to pipeline
            return;
        }

        /// <summary>
        /// Build a new subject, if the first 10 chars of the original subject are a valid date.
        /// We muste transform the de-DE format dd.MM.yyyy to yyyyMMdd for better sortability in the email client.
        /// </summary>
        /// <param name="MessageSubject">The original subject string</param>
        /// <returns>The modified subject string, if modification was possible</returns>
        private static string ModifiySubject(string MessageSubject)
        {
            string newSubject = String.Empty;

            if (MessageSubject.Length >= 10)
            {
                string dateCheck = MessageSubject.Substring(0, 10);
                DateTime dt = new DateTime();
                try
                {
                    // Check if we can parse the datetime
                    if (DateTime.TryParse(dateCheck, out dt))
                    {
                        // lets fetch the subject starting at the 10th character
                        string subjectRight = MessageSubject.Substring(10).Trim();
                        // build a new subject
                        newSubject = string.Format("{0:yyyyMMdd} {1}", dt, subjectRight);
                    }
                }
                finally
                {
                    // do nothing
                }
            }

            return newSubject;
        }


        /// <summary>
        /// Replace invalid filename chars with an underscore
        /// </summary>
        /// <param name="name">The filename to be checked</param>
        /// <returns>The sanitized filename</returns>
        private static string MakeValidFileName(string name)
        {
            string invalidChars = Regex.Escape(new string(Path.GetInvalidFileNameChars()));
            string invalidRegExStr = string.Format(@"[{0}]+", invalidChars);
            return Regex.Replace(name, invalidRegExStr, "_");
        }

    }

    #endregion

    #region Message Modifier Configuration

    /// <summary>
    /// Message Modifier Configuration class
    /// </summary>
    public class MessageModifierConfig
    {
        /// <summary>
        ///  The name of the configuration file.
        /// </summary>
        private static readonly string configFileName = "SFTools.MessageModify.Config.xml";

        /// <summary>
        /// Point out the directory with the configuration file (= assembly location)
        /// </summary>
        private string configDirectory;

        /// <summary>
        /// The filesystem watcher to monitor configuration file updates.
        /// </summary>
        private FileSystemWatcher configFileWatcher;

        /// <summary>
        /// The from address
        /// </summary>
        private Dictionary<string, RoutingAddress> addressMap;

        /// <summary>
        /// Whether reloading is ongoing
        /// </summary>
        private int reLoading = 0;

        /// <summary>
        /// The mapping between domain to catchall address.
        /// </summary>
        public Dictionary<string, RoutingAddress> AddressMap
        {
            get { return this.addressMap; }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        public MessageModifierConfig()
        {
            // Setup a file system watcher to monitor the configuration file
            this.configDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            this.configFileWatcher = new FileSystemWatcher(this.configDirectory);
            this.configFileWatcher.NotifyFilter = NotifyFilters.LastWrite;
            this.configFileWatcher.Filter = configFileName;
            this.configFileWatcher.Changed += new FileSystemEventHandler(this.OnChanged);

            // Create an initially empty map
            this.addressMap = new Dictionary<string, RoutingAddress>();

            // Load the configuration
            this.Load();

            // Now start monitoring
            this.configFileWatcher.EnableRaisingEvents = true;
        }

        /// <summary>
        /// Configuration changed handler.
        /// </summary>
        /// <param name="source">Event source.</param>
        /// <param name="e">Event arguments.</param>
        private void OnChanged(object source, FileSystemEventArgs e)
        {
            // Ignore if load ongoing
            if (Interlocked.CompareExchange(ref this.reLoading, 1, 0) != 0)
            {
                Trace.WriteLine("load ongoing: ignore");
                return;
            }

            // (Re) Load the configuration
            this.Load();

            // Reset the reload indicator
            this.reLoading = 0;
        }

        /// <summary>
        /// Load the configuration file. If any errors occur, does nothing.
        /// </summary>
        private void Load()
        {
            // Load the configuration
            XmlDocument doc = new XmlDocument();
            bool docLoaded = false;
            string fileName = Path.Combine(this.configDirectory, MessageModifierConfig.configFileName);

            try
            {
                doc.Load(fileName);
                docLoaded = true;
            }
            catch (FileNotFoundException)
            {
                Trace.WriteLine("Configuration file not found: {0}", fileName);
            }
            catch (XmlException e)
            {
                Trace.WriteLine("XML error: {0}", e.Message);
            }
            catch (IOException e)
            {
                Trace.WriteLine("IO error: {0}", e.Message);
            }

            // If a failure occured, ignore and simply return
            if (!docLoaded || doc.FirstChild == null)
            {
                Trace.WriteLine("Configuration error: either no file or an XML error");
                return;
            }

            // Create a dictionary to hold the mappings
            Dictionary<string, RoutingAddress> map = new Dictionary<string, RoutingAddress>(100);

            // Track whether there are invalid entries
            bool invalidEntries = false;

            // Validate all entries and load into a dictionary
            foreach (XmlNode node in doc.FirstChild.ChildNodes)
            {
                if (string.Compare(node.Name, "domain", true, CultureInfo.InvariantCulture) != 0)
                {
                    continue;
                }

                XmlAttribute domain = node.Attributes["name"];
                XmlAttribute address = node.Attributes["address"];

                // Validate the data
                if (domain == null || address == null)
                {
                    invalidEntries = true;
                    Trace.WriteLine("Reject configuration due to an incomplete entry. (Either or both domain and address missing.)");
                    break;
                }

                if (!RoutingAddress.IsValidAddress(address.Value))
                {
                    invalidEntries = true;
                    Trace.WriteLine(String.Format("Reject configuration due to an invalid address ({0}).", address));
                    break;
                }

                // Add the new entry
                string lowerDomain = domain.Value.ToLower();
                map[lowerDomain] = new RoutingAddress(address.Value);

                Trace.WriteLine(String.Format("Added entry ({0} -> {1})", lowerDomain, address.Value));
            }

            // If there are no invalid entries, swap in the map
            if (!invalidEntries)
            {
                Interlocked.Exchange<Dictionary<string, RoutingAddress>>(ref this.addressMap, map);
                Trace.WriteLine("Accepted configuration");
            }
        }
    }

    #endregion
}
