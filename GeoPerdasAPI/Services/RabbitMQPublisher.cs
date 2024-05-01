using GeoPerdasAPI.Config;
using GeoPerdasCloud.ProgGeoPerdas.Legacy.Config;
using Microsoft.AspNetCore.Connections;
using Newtonsoft.Json;
using RabbitMQ.Client;
using System.Text;
using System.Text.Json.Serialization;

namespace GeoPerdasAPI.Services
{
    public class RabbitMQPublisher
    {
        private const string QUEU_NAME = "GeoperdasCalculo";
        public void Publish(FormConfigControls message)
        {
            var config = new RabbitConfig();
            try
            {
                var factory = new ConnectionFactory()
                {
                    HostName = config.Server,
                    Port=Convert.ToInt32(config.Port),
                    UserName=config.User,
                    Password=config.Password,
                    VirtualHost=config.VirtualHost
                };

                using var connection = factory.CreateConnection();
                using var channel = connection.CreateModel();

                // Cria uma fila caso ela não exista
                channel.QueueDeclare(queue: QUEU_NAME,
                                     durable: true,
                                     exclusive: false,
                                     autoDelete: false,
                                     arguments: null);

                var body = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(message));

                // Publica a mensagem na fila
                channel.BasicPublish(exchange: "",
                                     routingKey: QUEU_NAME,
                                     basicProperties: null,
                                     body: body); ;

                Console.WriteLine($"Mensagem publicada: {message}");
            }
            catch (Exception) {
                throw new ArgumentException($"Erro ao conectar com o rabbit: {config.Server}:{config.Port}/{config.VirtualHost}  {config.User}:{config.Password}");
            }
        }
    }
}
