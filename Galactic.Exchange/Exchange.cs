using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.IO;
using Galactic.Configuration;
using System.Net;

namespace Galactic.Exchange
{
    public class Exchange
    {
        //The object representing the connection to the Exchange server.
        private ExchangeService service;

        /// <summary>
        /// Performs initial setup on ExchangeService object.
        /// </summary>
        /// <param name="configurationItemDirectoryPath">Path to the configuration items folder.</param>
        /// <param name="configurationItemName">Name of the configuration item.</param>
        public Exchange(string configurationItemDirectoryPath, string configurationItemName)
        {
            if (!string.IsNullOrWhiteSpace(configurationItemDirectoryPath) && !string.IsNullOrWhiteSpace(configurationItemName))
            {
                // Get the configuration item with the connection data from a file.
                ConfigurationItem configItem = new ConfigurationItem(configurationItemDirectoryPath, configurationItemName, true);

                // Get the connection data from the configuration item.
                StringReader reader = new StringReader(configItem.Value);

                //Read Exchange version from config file and setup ExchangeService.
                string exchangeVersion = reader.ReadLine();

                switch (exchangeVersion)
                {
                    case "Exchange2007_SP1":
                        service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                        break;
                    case "Exchange2010":
                        service = new ExchangeService(ExchangeVersion.Exchange2010);
                        break;
                    case "Exchange2010_SP1":
                        service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
                        break;
                    case "Exchange2010_SP2":
                        service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
                        break;
                    case "Exchange2013":
                        service = new ExchangeService(ExchangeVersion.Exchange2013);
                        break;
                    case "Exchange2013_SP1":
                        service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                        break;
                    default:
                        throw new Exception("Exchange version supplied does not match any known version. Please check configuration file for errors");
                }

                //Setup strings for the configuration item.
                string userName;
                string password;
                string domain;

                //Select connection type and bind the credentials to the ExchangeService.
                string connectionType = reader.ReadLine();
                string address = reader.ReadLine();
                switch (connectionType)
                {
                    case "autoDiscovery":
                        userName = reader.ReadLine();
                        password = reader.ReadLine();
                        if (!string.IsNullOrWhiteSpace(userName) && !string.IsNullOrWhiteSpace(password))
                        {
                            service.Credentials = new NetworkCredential(userName, password);
                            service.AutodiscoverUrl(address);
                        }
                        else
                        {
                            throw new Exception("Username or password null/empty string. Please check configuration item.");
                        }
                        break;
                    case "autoDiscoveryDomain":
                        userName = reader.ReadLine();
                        password = reader.ReadLine();
                        domain = reader.ReadLine();
                        if (!string.IsNullOrWhiteSpace(userName) && !string.IsNullOrWhiteSpace(password) && !string.IsNullOrWhiteSpace(domain))
                        {
                            service.Credentials = new NetworkCredential(userName, password, domain);
                            service.AutodiscoverUrl(address);
                        }
                        else
                        {
                            throw new Exception("Username, password, or domain null/empty string. Please check configuration item.");
                        }
                        break;
                    case "manualUrl":
                        userName = reader.ReadLine();
                        password = reader.ReadLine();
                        if (!string.IsNullOrWhiteSpace(userName) && !string.IsNullOrWhiteSpace(password))
                        {
                            service.Credentials = new NetworkCredential(userName, password);
                            service.Url = new Uri(address);
                        }
                        else
                        {
                            throw new Exception("Username or password null/empty string. Please check configuration item.");
                        }
                        break;
                    case "manualUrlDomain":
                        userName = reader.ReadLine();
                        password = reader.ReadLine();
                        domain = reader.ReadLine();
                        if (!string.IsNullOrWhiteSpace(userName) && !string.IsNullOrWhiteSpace(password) && !string.IsNullOrWhiteSpace(domain))
                        {
                            service.Credentials = new NetworkCredential(userName, password, domain);
                            service.Url = new Uri(address);
                        }
                        else
                        {
                            throw new Exception("Username, password, or domain null/empty string. Please check configuration item.");
                        }
                        break;
                    default:
                        throw new Exception("Invalid connection type. Please check configuration file.");
                }


            }
            else
            {
                throw new Exception("Configuration item not found.");
            }
        }

        /// <summary>
        /// Performs initial setup on ExchangeService object.
        /// </summary>
        /// <param name="exchangeVersion">The version of Exchange on the server.</param>
        /// <param name="address">The email address of the account or EWS URL.</param>
        /// <param name="username">Account username.</param>
        /// <param name="password">Account password.</param>
        /// <param name="isAutoDetect">Sets whether to autodetect url or set manually. True = Autodetect, False = Manual.</param>
        public Exchange(string exchangeVersion, string address, string userName, string password, bool isAutoDetect)
        {
            if (!string.IsNullOrWhiteSpace(exchangeVersion) && !string.IsNullOrWhiteSpace(address) && !string.IsNullOrWhiteSpace(userName) && !string.IsNullOrWhiteSpace(password))
            {
                //Set correct version of Exchange.
                switch (exchangeVersion)
                {
                    case "Exchange2007_SP1":
                        service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                        break;
                    case "Exchange2010":
                        service = new ExchangeService(ExchangeVersion.Exchange2010);
                        break;
                    case "Exchange2010_SP1":
                        service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
                        break;
                    case "Exchange2010_SP2":
                        service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
                        break;
                    case "Exchange2013":
                        service = new ExchangeService(ExchangeVersion.Exchange2013);
                        break;
                    case "Exchange2013_SP1":
                        service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                        break;
                    default:
                        throw new Exception("Exchange version supplied does not match any known version. Please check configuration file for errors");
                }

                if(isAutoDetect)
                {
                    service.Credentials = new NetworkCredential(userName, password);
                    service.AutodiscoverUrl(address);
                }
                else
                {
                    service.Credentials = new NetworkCredential(userName, password);
                    service.Url = new Uri(address);
                }
            }
            else
            {
                throw new Exception("One or more arguments are incorrect.");
            }
        }

        /// <summary>
        /// Performs initial setup on ExchangeService object.
        /// </summary>
        /// <param name="exchangeVersion">The version of Exchange on the server.</param>
        /// <param name="address">The email address of the account or EWS URL.</param>
        /// <param name="username">Account username.</param>
        /// <param name="password">Account password.</param>
        /// <param name="domain">The AD domain that the accound resides on.</param>
        /// <param name="isAutoDetect">Sets whether to autodetect url or set manually. True = Autodetect, False = Manual.</param>
        public Exchange(string exchangeVersion, string address, string userName, string password, string domain, bool isAutoDetect)
        {
            if (!string.IsNullOrWhiteSpace(exchangeVersion) && !string.IsNullOrWhiteSpace(address) && !string.IsNullOrWhiteSpace(userName) && !string.IsNullOrWhiteSpace(password) && !string.IsNullOrWhiteSpace(domain))
            {
                //Set correct version of Exchange.
                switch (exchangeVersion)
                {
                    case "Exchange2007_SP1":
                        service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                        break;
                    case "Exchange2010":
                        service = new ExchangeService(ExchangeVersion.Exchange2010);
                        break;
                    case "Exchange2010_SP1":
                        service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
                        break;
                    case "Exchange2010_SP2":
                        service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
                        break;
                    case "Exchange2013":
                        service = new ExchangeService(ExchangeVersion.Exchange2013);
                        break;
                    case "Exchange2013_SP1":
                        service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                        break;
                    default:
                        throw new Exception("Exchange version supplied does not match any known version. Please check configuration file for errors");
                }

                if (isAutoDetect)
                {
                    service.Credentials = new NetworkCredential(userName, password, domain);
                    service.AutodiscoverUrl(address);
                }
                else
                {
                    service.Credentials = new NetworkCredential(userName, password, domain);
                    service.Url = new Uri(address);
                }
            }
            else
            {
                throw new Exception("One or more arguments are incorrect.");
            }
        }
    }
}
