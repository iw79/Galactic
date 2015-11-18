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
        }
    }
}
