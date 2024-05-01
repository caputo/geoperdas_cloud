using GeoPerdasCloud.ProgGeoPerdas.Legacy.Config;

namespace GeoPerdasAPI.Services
{
    public class SequenceRequestService
    {
        public static bool RequestSequenceMessage(FormConfigControls config)
        {
            RabbitMQPublisher publisher = new RabbitMQPublisher();
            try
            {
                foreach (var feeder in config.tbObjectFeeder.Split(';'))
                {
                    var request = config.Clone();
                    request.tbObjectFeeder = feeder;
                    publisher.Publish(request);
                }
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }

}
