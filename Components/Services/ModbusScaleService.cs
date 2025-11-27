using FluentModbus;
using Microsoft.Extensions.Logging;

namespace FirstWeigh.Services
{
    public class ModbusScaleService : IDisposable
    {
        private readonly ILogger<ModbusScaleService> _logger;
        private ModbusTcpClient? _plcClient;

        private readonly string _plcIpAddress;
        private readonly byte _plcUnitId;
        private readonly ushort _scale1RegisterAddress;
        private readonly ushort _scale2RegisterAddress;

        public decimal Scale1Weight { get; private set; }
        public decimal Scale2Weight { get; private set; }

        private readonly List<decimal> _scale1History = new();
        private readonly List<decimal> _scale2History = new();
        private const int HISTORY_SIZE = 5;

        public bool Scale1Connected { get; private set; }
        public bool Scale2Connected { get; private set; }
        public bool IsConnected => _plcClient?.IsConnected ?? false;
        public string PlcIpAddress => _plcIpAddress;

        public event Action<decimal, decimal>? WeightUpdated;

        public ModbusScaleService(IConfiguration configuration, ILogger<ModbusScaleService> logger)
        {
            _logger = logger;
            _plcIpAddress = configuration["Modbus:PlcIpAddress"] ?? "192.168.1.100";
            _plcUnitId = byte.Parse(configuration["Modbus:PlcUnitId"] ?? "1");
            _scale1RegisterAddress = ushort.Parse(configuration["Modbus:Scale1RegisterAddress"] ?? "0");
            _scale2RegisterAddress = ushort.Parse(configuration["Modbus:Scale2RegisterAddress"] ?? "2");

            _logger.LogInformation("ModbusScaleService initialized - PLC: {PlcIp}, Unit: {UnitId}, Scale1 Reg: {Scale1Reg}, Scale2 Reg: {Scale2Reg}",
                _plcIpAddress, _plcUnitId, _scale1RegisterAddress, _scale2RegisterAddress);
        }

        public async Task<bool> ConnectAsync()
        {
            Scale1Connected = false;
            Scale2Connected = false;

            try
            {
                _plcClient = new ModbusTcpClient();
                _logger.LogInformation("Attempting to connect to PLC at {PlcIp}...", _plcIpAddress);

                _plcClient.Connect(_plcIpAddress, ModbusEndianness.BigEndian);
                _plcClient.ReadTimeout = 5000;
                _plcClient.WriteTimeout = 5000;

                if (_plcClient?.IsConnected ?? false)
                {
                    _logger.LogInformation("PLC connected successfully at {PlcIp}", _plcIpAddress);

                    try
                    {
                        await ReadScaleWeightAsync(_scale1RegisterAddress);
                        Scale1Connected = true;
                        _logger.LogInformation("Scale 1 connected - Register {Register}", _scale1RegisterAddress);
                    }
                    catch (Exception ex)
                    {
                        Scale1Connected = false;
                        _logger.LogError(ex, "Scale 1 connection failed - Register {Register}", _scale1RegisterAddress);
                    }

                    try
                    {
                        await ReadScaleWeightAsync(_scale2RegisterAddress);
                        Scale2Connected = true;
                        _logger.LogInformation("Scale 2 connected - Register {Register}", _scale2RegisterAddress);
                    }
                    catch (Exception ex)
                    {
                        Scale2Connected = false;
                        _logger.LogError(ex, "Scale 2 connection failed - Register {Register}", _scale2RegisterAddress);
                    }

                    return Scale1Connected || Scale2Connected;
                }

                _logger.LogWarning("PLC client created but not connected");
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to connect to PLC at {PlcIp}", _plcIpAddress);

                if (_plcClient != null)
                {
                    try
                    {
                        _plcClient.Disconnect();
                        _plcClient.Dispose();
                    }
                    catch { }
                    _plcClient = null;
                }

                Scale1Connected = false;
                Scale2Connected = false;
                return false;
            }
        }

        public async Task StartReadingAsync(CancellationToken cancellationToken = default)
        {
            if (_plcClient == null || !_plcClient.IsConnected)
            {
                throw new InvalidOperationException("Not connected to PLC");
            }

            _logger.LogInformation("Starting Modbus polling...");
            _scale1History.Clear();
            _scale2History.Clear();

            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    try
                    {
                        if (Scale1Connected)
                        {
                            try
                            {
                                Scale1Weight = await ReadScaleWeightAsync(_scale1RegisterAddress);
                                UpdateWeightHistory(_scale1History, Scale1Weight);
                            }
                            catch (Exception ex)
                            {
                                _logger.LogWarning(ex, "Error reading Scale 1");
                                Scale1Weight = 0;
                            }
                        }

                        if (Scale2Connected)
                        {
                            try
                            {
                                Scale2Weight = await ReadScaleWeightAsync(_scale2RegisterAddress);
                                UpdateWeightHistory(_scale2History, Scale2Weight);
                            }
                            catch (Exception ex)
                            {
                                _logger.LogWarning(ex, "Error reading Scale 2");
                                Scale2Weight = 0;
                            }
                        }

                        WeightUpdated?.Invoke(Scale1Weight, Scale2Weight);
                        await Task.Delay(500, cancellationToken);
                    }
                    catch (OperationCanceledException)
                    {
                        _logger.LogInformation("Modbus polling cancelled");
                        break;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Unexpected error in Modbus loop");
                        if (cancellationToken.IsCancellationRequested)
                            break;
                        await Task.Delay(1000, cancellationToken);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                _logger.LogInformation("Modbus polling stopped");
            }
            finally
            {
                _logger.LogInformation("Modbus polling loop exited cleanly");
            }
        }

        private async Task<decimal> ReadScaleWeightAsync(ushort registerAddress)
        {
            if (_plcClient == null || !_plcClient.IsConnected)
                return 0;

            try
            {
                var registers = _plcClient.ReadHoldingRegisters<ushort>(_plcUnitId, registerAddress, 2);

                var bytes = new byte[4];
                bytes[0] = (byte)(registers[1] & 0xFF);
                bytes[1] = (byte)((registers[1] >> 8) & 0xFF);
                bytes[2] = (byte)(registers[0] & 0xFF);
                bytes[3] = (byte)((registers[0] >> 8) & 0xFF);
                float weight = BitConverter.ToSingle(bytes, 0);

                return await Task.FromResult((decimal)weight);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Error reading register {Register}", registerAddress);
                return 0;
            }
        }

        public async Task TareScaleAsync(int scaleNumber)
        {
            if (_plcClient == null || !_plcClient.IsConnected)
            {
                _logger.LogWarning("Cannot tare scale - PLC not connected");
                return;
            }

            try
            {
                ushort tareRegister = scaleNumber == 1 ? (ushort)100 : (ushort)101;
                await Task.Run(() => _plcClient.WriteSingleRegister(_plcUnitId, tareRegister, 1));
                _logger.LogInformation("Tare command sent to Scale {ScaleNumber} - Register {Register}", scaleNumber, tareRegister);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error taring Scale {ScaleNumber}", scaleNumber);
            }
        }

        public bool IsScale1Stable(decimal tolerance = 0.005m)
        {
            if (_scale1History.Count < 3) return false;
            return (_scale1History.Max() - _scale1History.Min()) <= tolerance;
        }

        public bool IsScale2Stable(decimal tolerance = 0.005m)
        {
            if (_scale2History.Count < 3) return false;
            return (_scale2History.Max() - _scale2History.Min()) <= tolerance;
        }

        public List<decimal> GetScale2History() => new List<decimal>(_scale2History);
        public List<decimal> GetScale1History() => new List<decimal>(_scale1History);

        private void UpdateWeightHistory(List<decimal> history, decimal newWeight)
        {
            history.Add(newWeight);
            if (history.Count > HISTORY_SIZE)
                history.RemoveAt(0);
        }

        public void Disconnect()
        {
            _logger.LogInformation("Disconnecting from PLC...");
            Scale1Connected = false;
            Scale2Connected = false;
            _scale1History.Clear();
            _scale2History.Clear();

            if (_plcClient != null)
            {
                try
                {
                    _plcClient.Disconnect();
                    _plcClient.Dispose();
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Error during PLC disconnect");
                }
                _plcClient = null;
            }
            _logger.LogInformation("PLC disconnected");
        }

        public void Dispose() => Disconnect();
    }
}