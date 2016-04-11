﻿using System.Net.Sockets;
using System.Net;
using System.Text;
using UnityEngine;

namespace KSP_Chroma_Control
{
    [KSPAddon(KSPAddon.Startup.Instantly, true)]
    public class KSPChromaPlugin : MonoBehaviour
    {
        private static KSPChromaPlugin singleton;
        public static KSPChromaPlugin getInstance()
        {
            return singleton;
        }

        private ColorSchemes.ColorScheme Scheme;
        private UdpClient Client { get; set; }

        public KSPChromaPlugin()
        {
            singleton = this;
        }

        void Awake()
        {
            this.Scheme = new ColorSchemes.LogoScheme();
            Client = new UdpClient();
            IPEndPoint ep = new IPEndPoint(IPAddress.Parse("127.0.0.1"), 8888);
            this.Client.Connect(ep);
        }

        void Start()
        {
            Debug.Log("Sending Keyboard Layout");
            byte[] ToSend = Encoding.UTF8.GetBytes(this.Scheme.ToString());
            this.Client.Send(ToSend, ToSend.Length);
        }

        void Update()
        {

        }

        void FixedUpdate()
        {
           
        }

        void OnDestroy()
        {
            
        }
    }
}