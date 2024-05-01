using RabbitMQ.Client.Events;
using RabbitMQ.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Extensions.Hosting;
using System.Text;
using System.Threading.Tasks;
using GeoPerdasCloud.ProgGeoPerdas.Legacy.Config;
using GeoPerdasCloud.ProgGeoPerdas.Legacy.LegacyCode;
using Newtonsoft.Json;
using Microsoft.Extensions.Configuration;
using System.Configuration;

namespace GeoPerdasCloud.ProgGeoPerdasConsole
{
    public class RabbitMQHosterService : IHostedService
    {
        private IModel channel = null;
        private IConnection connection = null;
        private String? _rabbitMqServer;
        public RabbitMQHosterService() {
           

            
        }


        // Initiate RabbitMQ and start listening to an input queue
        private void Run()
        {
            var config = new RabbitConfig();   
          

            // ! Fill in your data here !
            var factory = new ConnectionFactory()
            {
                HostName = config.Server,
                Port = Convert.ToInt32(config.Port),
                VirtualHost = config.VirtualHost,
                UserName = config.User,
                Password = config.Password
            };

            this.connection = factory.CreateConnection();
            this.channel = this.connection.CreateModel();

            this.channel.QueueDeclare(queue: "GeoperdasCalculo",
                                durable: true,
                                exclusive: false,
                                autoDelete: false,
                                arguments: null);            


            this.channel.BasicQos(prefetchSize: 0, prefetchCount: 1, global: false);

            Console.WriteLine(" [*] Waiting for messages.");

            var consumer = new EventingBasicConsumer(this.channel);
            consumer.Received += OnMessageRecieved;

            this.channel.BasicConsume(queue: "GeoperdasCalculo",
                                autoAck: false,
                                consumer: consumer);
        }

        public Task StartAsync(CancellationToken cancellationToken)
        {
            this.Run();
            return Task.CompletedTask;
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            this.channel.Dispose();
            this.connection.Dispose();
            return Task.CompletedTask;
        }
        private void OnMessageRecieved(object model, BasicDeliverEventArgs args)
        {
            Console.WriteLine("Nova mensagem");
            try
            {
                var body = args.Body;
                var message = Encoding.UTF8.GetString(body.ToArray());
                var configCalc = JsonConvert.DeserializeObject<FormConfigControls>(message);          
                using var pgpForm = new ProgGeoperdasForm(configCalc);
                pgpForm.bConnection_Click();
                pgpForm.bExecuteBD_Click();
                Console.WriteLine(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("erro na mensagem");
                Console.Write(ex.Message);
            }
            finally {
                channel.BasicAck(args.DeliveryTag, true);
            }
        }
    }
}
