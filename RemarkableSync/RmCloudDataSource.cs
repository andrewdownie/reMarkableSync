﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Net;
using System.IO.Compression;
using System.Threading.Tasks;

// TODO: Exception handling
namespace RemarkableSync
{
    public class RmCloudDataSource : IRmDataSource
    {
        private static string DeviceTokenName = "rmdevicetoken";
        private static string UserTokenName = "rmusertoken";
        private static string EmptyToken = "****";
        private static string UserAgent = "rmapi";
        private static string Device = "desktop-windows";
        private static string DeviceTokenUrl = "https://webapp-production-dot-remarkable-production.appspot.com/token/json/2/device/new";
        private static string UserTokenUrl = "https://webapp-production-dot-remarkable-production.appspot.com/token/json/2/user/new";
        private static string BaseUrl = "https://document-storage-production-dot-remarkable-production.appspot.com"; 

        private string _devicetoken;
        private string _usertoken;
        private bool _initialized;

        private HttpClient _client;
        private IConfigStore _configStore;

        public RmCloudDataSource(IConfigStore configStore)
        {
            _usertoken = null;
            _devicetoken = null;
            _initialized = false;
            _configStore = configStore;

            _client = new HttpClient();
            _client.DefaultRequestHeaders.Add("user-agent", UserAgent);  
        }

        public async Task<bool> RegisterWithOneTimeCode(string oneTimeCode)
        {
            string uuid = Guid.NewGuid().ToString();
            string requestString = $@"{{
                ""code"": ""{oneTimeCode}"",
                ""deviceDesc"": ""{Device}"",
                ""deviceID"": ""{uuid}""
            }}";

            try
            {
                Console.WriteLine($"RmCloudDataSource::RegisterWithOneTimeCode() - registring with code: {oneTimeCode}");
                HttpResponseMessage response = await Request(
                    HttpMethod.Post,
                    DeviceTokenUrl,
                    null,
                    new ByteArrayContent(Encoding.ASCII.GetBytes(requestString)));

                if (response.IsSuccessStatusCode)
                {
                    byte[] responseContent = response.Content.ReadAsByteArrayAsync().Result;
                    _devicetoken = Encoding.ASCII.GetString(responseContent);
                    WriteConfig();
                    return true;
                }
                else
                {
                    Console.WriteLine($"RmCloudDataSource::RegisterWithOneTimeCode() - response code: {response.StatusCode}");
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("RmCloudDataSource::RegisterWithOneTimeCode() - Error: " + err.Message);
            }
            return false;
        }

        public async Task<List<RmItem>> GetItemHierarchy()
        {
            List<RmItem> collection = await GetAllItems();
            return getChildItemsRecursive("", ref collection);
        }

        private async Task<List<RmItem>> GetAllItems()
        {
            if (!_initialized)
            {
                await Initialize();
            }

            HttpResponseMessage response = await Request(
                HttpMethod.Get,
                "/document-storage/json/2/docs",
                null,
                null);

            if (!response.IsSuccessStatusCode)
            {
                string errMsg = "GetAllItems request failed with status code " + response.StatusCode.ToString();
                throw new Exception(errMsg);
            }

            string responseContent = await response.Content.ReadAsStringAsync();
            List<RmItem> collection = JsonSerializer.Deserialize<List<RmItem>>(responseContent);
            return collection;
        }

        public async Task<RmDownloadedDoc> DownloadDocument(RmItem item)
        {
            if (!_initialized)
            {
                await Initialize();
            }

            if (item.Type != RmItem.DocumentType)
            {
                Console.WriteLine($"RmCloudDataSource::DownloadDocument() - item with id {item.ID} is not document type");
                return null;

            }

            try
            {
                // first get the blob url
                string url = $"/document-storage/json/2/docs?doc={WebUtility.UrlEncode(item.ID)}&withBlob=true";
                HttpResponseMessage response = await Request(HttpMethod.Get, url, null, null);
                if (!response.IsSuccessStatusCode)
                {
                    Console.WriteLine("RmCloudDataSource::DownloadDocument() -  request failed with status code " + response.StatusCode.ToString());
                    return null;
                }
                List<RmItem> items = JsonSerializer.Deserialize<List<RmItem>>(response.Content.ReadAsStringAsync().Result);
                if (items.Count == 0)
                {
                    Console.WriteLine("RmCloudDataSource::DownloadDocument() - Failed to find document with id: " + item.ID);
                    return null;
                }
                string blobUrl = items[0].BlobURLGet;
                Stream stream = await _client.GetStreamAsync(blobUrl);
                ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read);

                return new RmCloudDownloadedDoc(archive, item.ID);
            }
            catch (Exception err)
            {
                Console.WriteLine($"RmCloud::DownloadDocument() - failed for id {item.ID}. Error: {err.Message}");
                return null;
            }

        }

        public void Dispose()
        {
            _client?.Dispose();
            _configStore?.Dispose();
        }

        private async Task<bool> Initialize()
        {
            LoadConfig();
            if (_devicetoken != null)
            {
                Console.WriteLine("device token loaded from config file");
                _initialized = await RenewToken();
            }
            return _initialized;
        }

        private void LoadConfig()
        {
            string devicetoken = _configStore.GetConfig(DeviceTokenName);
            string usertoken = _configStore.GetConfig(UserTokenName);
            _devicetoken = devicetoken == EmptyToken ? "" : devicetoken;
            _usertoken = usertoken == EmptyToken ? "" : usertoken;
        }

        private void WriteConfig()
        {
            Dictionary<string, string> mapConfigs = new Dictionary<string, string>();
            mapConfigs[DeviceTokenName] = _devicetoken?.Length > 0 ? _devicetoken : EmptyToken;
            mapConfigs[UserTokenName] = _usertoken?.Length > 0 ? _usertoken : EmptyToken;
            _configStore.SetConfigs(mapConfigs);

            _client.DefaultRequestHeaders.Add("Authorization", $"Bearer {_usertoken}");
        }

        private async Task<HttpResponseMessage> Request(HttpMethod method, string url, Dictionary<string, string> header, HttpContent content)
        {
            if (!url.StartsWith("http"))
            {
                if (!url.StartsWith("/"))
                    url = "/" + url;
                url = BaseUrl + url;
            }

            Console.WriteLine($"Request() -  url is: {url}");
            var request = new HttpRequestMessage();
            request.RequestUri = new Uri(url);
            request.Method = method;
            if (content != null)
            {
                request.Content = content;
            }

            // add/replace the supplied headers
            if (header != null)
            {
                foreach (var key in header.Keys)
                {
                    request.Headers.Add(key, header[key]);
                }
            }

            HttpResponseMessage response = await _client.SendAsync(request);
            return response;
        }

        private async Task<bool> RenewToken()
        {
            if (_devicetoken == null || _devicetoken.Length == 0)
            {
                throw new Exception("Please register a device first");
            }

            Dictionary<string, string> header = new Dictionary<string, string>();
            header.Add("Authorization", $"Bearer {_devicetoken}");

            HttpResponseMessage response = await Request(HttpMethod.Post, UserTokenUrl, header, null);
            if (response.IsSuccessStatusCode)
            {
                byte[] responseContent = response.Content.ReadAsByteArrayAsync().Result;
                _usertoken = Encoding.ASCII.GetString(responseContent);
                WriteConfig();
                Console.WriteLine("user token renewed");
                return true;
            }
            else
            {
                throw new Exception("Can't register device");
            }
        }

        private bool IsAuthenticated()
        {
            return _devicetoken != null && _devicetoken.Length > 0 && _usertoken != null && _usertoken.Length > 0;
        }

        private List<RmItem> getChildItemsRecursive(string parentId, ref List<RmItem> items)
        {
            var children = (from item in items where item.Parent == parentId select item).ToList(); 
            foreach(var child in children)
            {
                child.Children = getChildItemsRecursive(child.ID, ref items);
            }
            return children;
        }
    }
}
