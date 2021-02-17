// BarAction.cs: Actions performed by bar items.
//
// Copyright 2020 Raising the Floor - International
//
// Licensed under the New BSD license. You may not use this file except in
// compliance with this License.
//
// You may obtain a copy of the License at
// https://github.com/GPII/universal/blob/master/LICENSE.txt

namespace Morphic.Client.Bar.Data.Actions
{
    using CountlySDK;
    using Microsoft.Extensions.Logging;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using System.Net.WebSockets;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using System.Windows.Media;

    /// <summary>
    /// An action for a bar item.
    /// </summary>
    [JsonObject(MemberSerialization.OptIn)]
    [JsonConverter(typeof(TypedJsonConverter), "kind", "shellExec")]
    public abstract class BarAction
    {
        [JsonProperty("identifier")]
        public string Id { get; set; } = string.Empty;

        /// <summary>
        /// Called by <c>Invoke</c> to perform the implementation-specific action invocation.
        /// </summary>
        /// <param name="source">Button ID, for multi-button bar items.</param>
        /// <param name="toggleState">New state, if the button is a toggle.</param>
        /// <returns></returns>
        protected abstract Task<bool> InvokeAsyncImpl(string? source = null, bool? toggleState = null);

        /// <summary>
        /// Invokes the action.
        /// </summary>
        /// <param name="source">Button ID, for multi-button bar items.</param>
        /// <param name="toggleState">New state, if the button is a toggle.</param>
        /// <returns></returns>
        public async Task<bool> InvokeAsync(string? source = null, bool? toggleState = null)
        {
            bool result;
            try
            {
                try
                {
                    result = await this.InvokeAsyncImpl(source, toggleState);
                }
                catch (Exception e) when (!(e is ActionException || e is OutOfMemoryException))
                {
                    throw new ActionException(e.Message, e);
                }
            }
            catch (ActionException e)
            {
                App.Current.Logger.LogError(e, $"Error while invoking action for bar {this.Id} {this}");

                if (e.UserMessage != null)
                {
                    MessageBox.Show($"There was a problem performing the action:\n\n{e.UserMessage}",
                        "Custom MorphicBar", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                result = false;
            }
            finally
            {
                // record telemetry data for this action
                await this.SendTelemetryForBarAction(source, toggleState);
            }

            return result;
        }

        // NOTE: we should refactor this functionality to functions attached to each button (similar to how action callbacks are invoked)
        private async Task SendTelemetryForBarAction(string? source = null, bool? toggleState = null) 
        {
            // handle actions which must be filted by id
            switch (this.Id)
            {
                case "magnify":
                    {
                        if (source == "on") 
                        {
                            await Countly.RecordEvent("magnifierShow");
                        }
                        else if (source == "off")
                        {
                            await Countly.RecordEvent("magnifierHide");
                        }
                    }
                    break;
                case "read-aloud":
                    {
                        if (source == "play")
                        {
                            await Countly.RecordEvent("readSelectedPlay");
                        } 
                        else if (source == "stop")
                        {
                            await Countly.RecordEvent("readSelectedStop");
                            break;
                        }
                    }
                    break;
                case "":
                    switch (source)
                    {
                        case "com.microsoft.windows.colorFilters/enabled":
                            {
                                if (toggleState == true)
                                {
                                    await Countly.RecordEvent("colorFiltersOn");
                                    return;
                                }
                                else
                                {
                                    await Countly.RecordEvent("colorFiltersOff");
                                    return;
                                }
                            }
                            break;
                        case "com.microsoft.windows.highContrast/enabled":
                            {
                                if (toggleState == true)
                                {
                                    await Countly.RecordEvent("highContrastOn");
                                    return;
                                }
                                else
                                {
                                    await Countly.RecordEvent("highContrastOff");
                                    return;
                                }
                            }
                            break;
                        case "com.microsoft.windows.nightMode/enabled":
                            {
                                if (toggleState == true)
                                {
                                    await Countly.RecordEvent("nightModeOn");
                                    return;
                                }
                                else
                                {
                                    await Countly.RecordEvent("nightModeOff");
                                    return;
                                }
                            }
                            break;
                        case "copy":
                            {
                                await Countly.RecordEvent("screenSnip");
                            }
                            break;
                        case "dark-mode":
                            {
                                if (toggleState == true)
                                {
                                    await Countly.RecordEvent("darkModeOn");
                                }
                                else
                                {
                                    await Countly.RecordEvent("darkModeOff");
                                }
                            }
                            break;
                        case null:
                            // no tags; this is the Morphie button or another custom element with no known tags
                            break;
                        default:
                            // we do not understand this action type (for telemetry logging purposes)
                            Debug.Assert(false, "Unknown Action ID (missing telemetry hooks)");
                            break;
                    }
                    break;
                case "screen-zoom":
                    // this action type's telemetry is logged elsewhere
                    break;
                default:
                    // we do not understand this action type (for telemetry logging purposes)
                    Debug.Assert(false, "Unknown Action ID (missing telemetry hooks)");
                    break;
            }
        }

        /// <summary>
        /// Resolves "{identifiers}" in a string with its value.
        /// </summary>
        /// <param name="arg"></param>
        /// <param name="source"></param>
        /// <returns>null if arg is null</returns>
        protected string? ResolveString(string? arg, string? source)
        {
            // Today, there is only "{button}".
            return arg?.Replace("{button}", source ?? string.Empty);
        }

        public virtual Uri? DefaultImageUri { get; }
        public virtual ImageSource? DefaultImageSource { get; }
        public virtual bool IsAvailable { get; protected set; } = true;

        public virtual void Deserialized(BarData barData)
        {
        }
    }

    [JsonTypeName("null")]
    public class NoOpAction : BarAction
    {
        protected override Task<bool> InvokeAsyncImpl(string? source = null, bool? toggleState = null)
        {
            return Task.FromResult(true);
        }
    }

    [JsonTypeName("internal")]
    public class InternalAction : BarAction
    {
        [JsonProperty("function", Required = Required.Always)]
        public string? FunctionName { get; set; }

        [JsonProperty("args")]
        public Dictionary<string, string> Arguments { get; set; } = new Dictionary<string, string>();

        public string? TelemetryEventName { get; set; }
        
        protected override Task<bool> InvokeAsyncImpl(string? source = null, bool? toggleState = null)
        {
            try
            {
                if (this.FunctionName == null)
                {
                    return Task.FromResult(true);
                }

                Dictionary<string, string> resolvedArgs = this.Arguments
                    .ToDictionary(kv => kv.Key, kv => this.ResolveString(kv.Value, source) ?? string.Empty);

                resolvedArgs.Add("state", toggleState == true ? "on" : "off");

                return InternalFunctions.Default.InvokeFunction(this.FunctionName, resolvedArgs);
            }
            finally
            {
                if (this.TelemetryEventName != null) 
                {
                    Countly.RecordEvent(this.TelemetryEventName!);
                }
            }
        }
    }

    [JsonTypeName("gpii")]
    public class GpiiAction : BarAction
    {
        [JsonProperty("data", Required = Required.Always)]
        public JObject RequestObject { get; set; } = null!;

        protected override async Task<bool> InvokeAsyncImpl(string? source = null, bool? toggleState = null)
        {
            ClientWebSocket socket = new ClientWebSocket();
            CancellationTokenSource cancel = new CancellationTokenSource();
            await socket.ConnectAsync(new Uri("ws://localhost:8081/pspChannel"), cancel.Token);

            string requestString = this.RequestObject.ToString();
            byte[] bytes = Encoding.UTF8.GetBytes(requestString);
            
            ArraySegment<byte> sendBuffer = new ArraySegment<byte>(bytes);
            await socket.SendAsync(sendBuffer, WebSocketMessageType.Text, true, cancel.Token);

            return true;
        }
    }

    [JsonTypeName("shellExec")]
    public class ShellExecuteAction : BarAction
    {
        [JsonProperty("run")]
        public string? ShellCommand { get; set; }

        protected override Task<bool> InvokeAsyncImpl(string? source = null, bool? toggleState = null)
        {
            bool success = true;
            if (!string.IsNullOrEmpty(this.ShellCommand))
            {
                Process? process = Process.Start(new ProcessStartInfo()
                {
                    FileName = this.ResolveString(this.ShellCommand, source),
                    UseShellExecute = true
                });
                success = process != null;
            }

            return Task.FromResult(success);
        }

        public override void Deserialized(BarData barData)
        {
        }
    }

    /// <summary>
    /// Exception that gets thrown by action invokers.
    /// </summary>
    public class ActionException : ApplicationException
    {
        /// <summary>
        /// The message displayed to the user. null to not display a message.
        /// </summary>
        public string? UserMessage { get; set; }

        public ActionException(string? userMessage)
            : this(userMessage, userMessage, null)
        {
        }
        public ActionException(string? userMessage, Exception innerException)
            : this(userMessage, userMessage, innerException)
        {
        }

        public ActionException(string? userMessage, string? internalMessage = null, Exception? innerException = null)
            : base(internalMessage ?? userMessage ?? innerException?.Message, innerException)
        {
            this.UserMessage = userMessage;
        }
    }

}