using System;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Net.Http;
using System.Text;
using Microsoft.CSharp;
using System.CodeDom.Compiler;
using System.Reflection;
using System.IO;
using Newtonsoft.Json.Linq;

namespace InventorChatGPT
{
    public partial class MyForm : Form
    {
        const string API_KEY = "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";

        public MyForm()
        {
            InitializeComponent();
        }

        private async void submitButton_Click(object sender, EventArgs e)
        {
            try
            {
                string request = inputTextBox.Text;
                if (request.Contains("\n"))
                    request = request.Replace("\n", " ");

                string invApp = File.ReadAllText("InvApp.txt");
                request = "Give me a standalone 'public static void Main()' method in C# to " + request + " in current Inventor window.You use this as Inventor.Application object " + invApp + ". Ignore code descriptions, comments and instructions.";

                string endpoint = "https://api.openai.com/v1/chat/completions";
                var messages = new[]
                {
                    new {role = "user", content = request}
                };

                var data = new
                {
                    model = "gpt-4-0613",
                    messages = messages,
                    temperature = 0
                };

                string jsonString = JsonConvert.SerializeObject(data);
                var content = new StringContent(jsonString, Encoding.UTF8, "application/json");

                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Add("authorization", "Bearer " + API_KEY);
                var response2 = await client.PostAsync(endpoint, content);
                string responseContent = await response2.Content.ReadAsStringAsync();
                var jsonResponse = JObject.Parse(responseContent);
                var response = jsonResponse["choices"][0]["message"]["content"].Value<string>();

                string input = response;
                if (input.IndexOf("```csharp") != -1)
                    input = input.Substring(input.IndexOf("```csharp") + 10);
                else if (input.IndexOf("```") != 1)
                    input = input.Substring(input.IndexOf("```") + 4);

                if (input.IndexOf("```") != -1)
                    input = input.Substring(0, input.IndexOf("```"));

                if (!input.Contains("using Inventor;\n"))
                    input = "using Inventor;\n" + input;

                if (!input.Contains("using System;\n"))
                    input = "using System;\n" + input;

                if (input.Contains("Document newDoc = invApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, false);"))
                    input = input.Replace("Document newDoc = invApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, false);", "Document newDoc = invApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject);");

                if (!input.Contains("public static void Main()") && input.Contains("static void Main()"))
                    input = input.Replace("static void Main()", "public static void Main()");

                else if (!input.Contains("public static void Main(string[] args)") && input.Contains("static void Main(string[] args)"))
                    input = input.Replace("static void Main(string[] args)", "public static void Main()");

                StreamWriter sw = new StreamWriter("Input.txt");
                outputTextBox.Text = input;
                sw.WriteLine(input);
                sw.Close();

                CSharpCodeProvider provider = new CSharpCodeProvider();
                CompilerParameters parameters = new CompilerParameters
                {
                    GenerateExecutable = false,
                    GenerateInMemory = true
                };

                parameters.ReferencedAssemblies.Add(@"C:\Program Files\Autodesk\Inventor 2023\Bin\Public Assemblies\Autodesk.Inventor.Interop.dll");
                parameters.ReferencedAssemblies.Add(@"C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.7.2\System.dll");
                parameters.ReferencedAssemblies.Add(@"C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.7.2\System.Windows.Forms.dll");
                CompilerResults results = provider.CompileAssemblyFromSource(parameters, input);

                if (results.Errors.HasErrors)
                {
                    Console.WriteLine("Errors building the code:");
                    foreach (CompilerError error in results.Errors)
                    {
                        Console.WriteLine(error.ToString());
                        outputTextBox.Text += "\n" + error.ToString();
                    }
                }
                else
                {
                    Assembly assembly = results.CompiledAssembly;
                    Type[] types = assembly.GetTypes();
                    MethodInfo method = types[0].GetMethod("Main");
                    method.Invoke(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}