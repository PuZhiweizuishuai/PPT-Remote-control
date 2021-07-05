using Microsoft.VisualBasic.Devices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace PhoneAsPrompter
{
    class GetIPUtil
    {
        public static string IPV4()
        {
            string ipv4 = GetLocalIPv4(NetworkInterfaceType.Wireless80211);
            if (ipv4 == "")
            {
                ipv4 = GetLocalIPv4(NetworkInterfaceType.Ethernet);
                if (ipv4 == "")
                {
                    ipv4 = GetLoacalIPMaybeVirtualNetwork();
                }
             }
            return ipv4;
        }

        private static string GetLoacalIPMaybeVirtualNetwork()
        {
            string name = Dns.GetHostName();
            IPAddress[] ipadrlist = Dns.GetHostAddresses(name);
            foreach (IPAddress ipa in ipadrlist)
            {
                if (ipa.AddressFamily == AddressFamily.InterNetwork)
                {
                    return ipa.ToString();
                }
            }
            return "没有连接网络，请链接网络后重试！";
        }

        public static string GetLocalIPv4(NetworkInterfaceType _type)
        {
            string output = "";
            foreach (NetworkInterface item in NetworkInterface.GetAllNetworkInterfaces())
            {
                //Console.WriteLine(item.NetworkInterfaceType.ToString());
                if (item.NetworkInterfaceType == _type && item.OperationalStatus == OperationalStatus.Up)
                {
                    foreach (UnicastIPAddressInformation ip in item.GetIPProperties().UnicastAddresses)
                    {
                        if (ip.Address.AddressFamily == AddressFamily.InterNetwork)
                        {
                            output = ip.Address.ToString();
                        }
                    }
                }
            }
            return output;
        }
    }
}
