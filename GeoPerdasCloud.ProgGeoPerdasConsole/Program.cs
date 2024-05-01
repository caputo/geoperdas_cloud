using GeoPerdasCloud.ProgGeoPerdasConsole;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

await new HostBuilder()
              .ConfigureServices((hostContext, services) =>
              {
                  services.AddHostedService<RabbitMQHosterService>(); // register our service here            
              })
             .RunConsoleAsync();




