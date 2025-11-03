using FluentModbus;
using Microsoft.Win32;

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

                // ✅ Connect directly without Task.Run wrapper
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
                // ✅ This will now catch the connection timeout exception
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
                        // ✅ Expected - cancellation requested, exit cleanly
                        Console.WriteLine("⏹️ Modbus polling cancelled");
                        break;
                    }
                    catch (Exception ex)
                    {
                        // ⚠️ Unexpected error - log and retry
                        Console.WriteLine($"❌ Unexpected error in Modbus loop: {ex.Message}");

                        // Check if cancelled before retrying
                        if (cancellationToken.IsCancellationRequested)
                            break;

                        await Task.Delay(1000, cancellationToken);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // Normal cancellation at outer level
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

                // Convert to weight - OPTION 1: 32-bit IEEE 754 Float (Most common)
                var bytes = new byte[4];
                bytes[0] = (byte)(registers[1] & 0xFF);
                bytes[1] = (byte)((registers[1] >> 8) & 0xFF);
                bytes[2] = (byte)(registers[0] & 0xFF);
                bytes[3] = (byte)((registers[0] >> 8) & 0xFF);
                float weight = BitConverter.ToSingle(bytes, 0);

                return await Task.FromResult((decimal)weight);

                // OPTION 2: 16-bit integer (scaled) - Uncomment if your PLC uses this format
                // int rawWeight = registers[0];
                // return await Task.FromResult(rawWeight / 1000m); // If weight is in grams

                // OPTION 3: 32-bit integer - Uncomment if your PLC uses this format
                // int rawWeight = (registers[0] << 16) | registers[1];
                // return await Task.FromResult(rawWeight / 1000m);
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
                // This is just an example - you need to check your PLC's Modbus map

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

        public void Disconnect()
        {
            Scale1Connected = false;
            Scale2Connected = false;

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