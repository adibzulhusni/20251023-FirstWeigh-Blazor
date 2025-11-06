using FluentModbus;

namespace FirstWeigh.Services
{
    public class ModbusScaleService : IDisposable
    {
        private ModbusTcpClient? _plcClient;

        // PLC Configuration
        private readonly string _plcIpAddress;
        private readonly byte _plcUnitId;

        // Register addresses for each scale
        private readonly ushort _scale1RegisterAddress;
        private readonly ushort _scale2RegisterAddress;

        // Current weights
        public decimal Scale1Weight { get; private set; }
        public decimal Scale2Weight { get; private set; }

        // ✅ NEW: Weight history for stability checking
        private readonly List<decimal> _scale1History = new();
        private readonly List<decimal> _scale2History = new();
        private const int HISTORY_SIZE = 5;

        // Connection status
        public bool Scale1Connected { get; private set; }
        public bool Scale2Connected { get; private set; }
        public bool IsConnected => _plcClient?.IsConnected ?? false;

        public string PlcIpAddress => _plcIpAddress;

        // Events
        public event Action<decimal, decimal>? WeightUpdated;

        public ModbusScaleService(IConfiguration configuration)
        {
            // PLC Connection
            _plcIpAddress = configuration["Modbus:PlcIpAddress"] ?? "192.168.1.100";
            _plcUnitId = byte.Parse(configuration["Modbus:PlcUnitId"] ?? "1");

            // Scale register addresses
            _scale1RegisterAddress = ushort.Parse(configuration["Modbus:Scale1RegisterAddress"] ?? "0");
            _scale2RegisterAddress = ushort.Parse(configuration["Modbus:Scale2RegisterAddress"] ?? "2");
        }

        public async Task<bool> ConnectAsync()
        {
            Scale1Connected = false;
            Scale2Connected = false;

            try
            {
                _plcClient = new ModbusTcpClient();

                Console.WriteLine($"🔌 Attempting to connect to PLC at {_plcIpAddress}...");

                _plcClient.Connect(_plcIpAddress, ModbusEndianness.BigEndian);
                _plcClient.ReadTimeout = 5000;
                _plcClient.WriteTimeout = 5000;

                if (_plcClient?.IsConnected ?? false)
                {
                    Console.WriteLine($"✅ PLC ({_plcIpAddress}): Connected");

                    // Test read both scales to verify connection
                    try
                    {
                        await ReadScaleWeightAsync(_scale1RegisterAddress);
                        Scale1Connected = true;
                        Console.WriteLine($"✅ Scale 1 (Register {_scale1RegisterAddress}): Connected");
                    }
                    catch (Exception ex)
                    {
                        Scale1Connected = false;
                        Console.WriteLine($"❌ Scale 1 (Register {_scale1RegisterAddress}): Failed - {ex.Message}");
                    }

                    try
                    {
                        await ReadScaleWeightAsync(_scale2RegisterAddress);
                        Scale2Connected = true;
                        Console.WriteLine($"✅ Scale 2 (Register {_scale2RegisterAddress}): Connected");
                    }
                    catch (Exception ex)
                    {
                        Scale2Connected = false;
                        Console.WriteLine($"❌ Scale 2 (Register {_scale2RegisterAddress}): Failed - {ex.Message}");
                    }

                    return Scale1Connected || Scale2Connected;
                }

                Console.WriteLine($"⚠️ PLC client created but not connected");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Failed to connect to PLC at {_plcIpAddress}");
                Console.WriteLine($"   Error: {ex.Message}");

                // Clean up
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

            Console.WriteLine("🔄 Starting Modbus polling...");

            // ✅ Clear history at start
            _scale1History.Clear();
            _scale2History.Clear();

            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    try
                    {
                        // Read Scale 1
                        if (Scale1Connected)
                        {
                            try
                            {
                                Scale1Weight = await ReadScaleWeightAsync(_scale1RegisterAddress);
                                UpdateWeightHistory(_scale1History, Scale1Weight);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error reading Scale 1: {ex.Message}");
                                Scale1Weight = 0;
                            }
                        }

                        // Read Scale 2
                        if (Scale2Connected)
                        {
                            try
                            {
                                Scale2Weight = await ReadScaleWeightAsync(_scale2RegisterAddress);
                                UpdateWeightHistory(_scale2History, Scale2Weight);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error reading Scale 2: {ex.Message}");
                                Scale2Weight = 0;
                            }
                        }

                        // Notify listeners
                        WeightUpdated?.Invoke(Scale1Weight, Scale2Weight);

                        // Wait before next reading (500ms = 2 readings per second)
                        await Task.Delay(500, cancellationToken);
                    }
                    catch (OperationCanceledException)
                    {
                        Console.WriteLine("⏹️ Modbus polling cancelled");
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Unexpected error in Modbus loop: {ex.Message}");

                        if (cancellationToken.IsCancellationRequested)
                            break;

                        await Task.Delay(1000, cancellationToken);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("⏹️ Modbus polling stopped");
            }
            finally
            {
                Console.WriteLine("✅ Modbus polling loop exited cleanly");
            }
        }

        private async Task<decimal> ReadScaleWeightAsync(ushort registerAddress)
        {
            if (_plcClient == null || !_plcClient.IsConnected)
                return 0;

            try
            {
                // Read holding registers (function code 03)
                // Read 2 registers for 32-bit float
                var registers = _plcClient.ReadHoldingRegisters<ushort>(
                    _plcUnitId,
                    registerAddress,
                    2
                );

                // Convert to weight - 32-bit IEEE 754 Float (Most common)
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
                Console.WriteLine($"Error reading register {registerAddress}: {ex.Message}");
                return 0;
            }
        }

        public async Task TareScaleAsync(int scaleNumber)
        {
            if (_plcClient == null || !_plcClient.IsConnected)
                return;

            try
            {
                // Send tare command to PLC
                // IMPORTANT: Update these register addresses based on your PLC configuration
                ushort tareRegister = scaleNumber == 1 ? (ushort)100 : (ushort)101;

                await Task.Run(() =>
                    _plcClient.WriteSingleRegister(_plcUnitId, tareRegister, 1)
                );

                Console.WriteLine($"Tare command sent to Scale {scaleNumber} (Register {tareRegister})");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error taring scale {scaleNumber}: {ex.Message}");
            }
        }

        // ✅ NEW: Check if Scale 1 is stable
        public bool IsScale1Stable(decimal tolerance = 0.005m)
        {
            if (_scale1History.Count < 3) return false;

            var max = _scale1History.Max();
            var min = _scale1History.Min();
            var range = max - min;

            return range <= tolerance;
        }

        // ✅ NEW: Check if Scale 2 is stable
        public bool IsScale2Stable(decimal tolerance = 0.005m)
        {
            if (_scale2History.Count < 3) return false;

            var max = _scale2History.Max();
            var min = _scale2History.Min();
            var range = max - min;

            return range <= tolerance;
        }

        // ✅ NEW: Get Scale 2 history for external stability checks
        public List<decimal> GetScale2History()
        {
            return new List<decimal>(_scale2History);
        }

        // ✅ NEW: Get Scale 1 history
        public List<decimal> GetScale1History()
        {
            return new List<decimal>(_scale1History);
        }

        // ✅ PRIVATE: Update weight history
        private void UpdateWeightHistory(List<decimal> history, decimal newWeight)
        {
            history.Add(newWeight);
            if (history.Count > HISTORY_SIZE)
            {
                history.RemoveAt(0);
            }
        }

        public void Disconnect()
        {
            Scale1Connected = false;
            Scale2Connected = false;

            // Clear history
            _scale1History.Clear();
            _scale2History.Clear();

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
        }

        public void Dispose()
        {
            Disconnect();
        }
    }
}