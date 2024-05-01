using Microsoft.Extensions.Configuration;
using RabbitMQ.Client.Events;
using RabbitMQ.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Extensions.Hosting;
using System.Text;

namespace GeoPerdasCloud.ProgGeoPerdasConsole
{
    public class RabbitConfig
    {
        private IConfiguration configuration;

        public RabbitConfig()
        {
            configuration = new ConfigurationBuilder()
                    .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                    .Build();
        }

        public string Server
        {
            get
            {
                return Environment.GetEnvironmentVariable("APPSETTINGS_RABBITMQSERVER")
                                          ?? configuration.GetValue<string>("RabbitMQ:Server") ?? "rabbitmqserver";
            }
        }

        public string Port
        {
            get
            {
                return Environment.GetEnvironmentVariable("APPSETTINGS_RABBITMQPORT")
                                          ?? configuration.GetValue<string>("RabbitMQ:Port") ?? "5672";
            }
        }

        public string User
        {
            get
            {
                return Environment.GetEnvironmentVariable("APPSETTINGS_RABBITMQUSER")
                                          ?? configuration.GetValue<string>("RabbitMQ:User") ?? "guest";
            }
        }

        public string Password
        {
            get
            {
                return Environment.GetEnvironmentVariable("APPSETTINGS_RABBITMQPASSWORD")
                                          ?? configuration.GetValue<string>("RabbitMQ:Password") ?? "guest";
            }
        }

        public string VirtualHost
        {
            get
            {
                return Environment.GetEnvironmentVariable("APPSETTINGS_RABBITMQVIRTUALHOST")
                                          ?? configuration.GetValue<string>("RabbitMQ:VirtualHost") ?? "/";
            }
        }


    }
}
