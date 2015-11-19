using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace Galactic.Exchange
{
    class EmailItem
    {
        private Exchange exchange;

        private GetItemResponse mail;

        public EmailItem(Exchange exchange, ItemId itemId)
        {

        }

        public EmailItem(Exchange exchange, GetItemResponse mail)
        {

        }
    }
}
