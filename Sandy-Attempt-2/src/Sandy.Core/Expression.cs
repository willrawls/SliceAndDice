using System.Globalization;
using System.Text.RegularExpressions;

namespace Sandy.Core;

/// <summary>Small, deterministic expression evaluator; never executes C# or VB source.</summary>
public static class SandyExpression
{
    public static string Evaluate(string expression, IDictionary<string, string> variables)
        => new Parser(expression, variables).Parse();

    internal static bool Truth(string value) => value.Equals("True",StringComparison.OrdinalIgnoreCase)||double.TryParse(value,NumberStyles.Float,CultureInfo.InvariantCulture,out var number)&&number!=0&&!double.IsNaN(number);
    internal static string Format(double value) => value.ToString("G", CultureInfo.InvariantCulture);
    internal static double Number(string value) => double.Parse(value, NumberStyles.Float, CultureInfo.InvariantCulture);

    private sealed class Parser
    {
        private static readonly Regex Tokenizer = new(@"\G\s*(""(?:[^""]|"""")*""|'[^']*'|&[Hh][0-9A-Fa-f]+|\[[^\]]+\]|(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?|<>|!=|==|<=|>=|[()+\-*/\\^<>=&]|[A-Za-z_][A-Za-z_0-9.]*)", RegexOptions.Compiled);
        private readonly List<string> tokens = new();
        private readonly IDictionary<string, string> variables;
        private int index;

        public Parser(string text, IDictionary<string, string> variables)
        {
            this.variables = variables;
            if (text.Length > 65536) throw new FormatException("Expression size limit exceeded.");
            var offset = 0;
            while (offset < text.Length && text.Substring(offset).Trim().Length > 0)
            {
                var match = Tokenizer.Match(text, offset);
                if (!match.Success) throw new FormatException($"Invalid expression near '{text.Substring(offset)}'.");
                var token=match.Groups[1].Value;token=token=="=="?"=":token=="!="?"<>":token;
                if(tokens.Count>0&&Bare(tokens[^1])&&Bare(token))tokens[^1]+=" "+token;else tokens.Add(token);
                if (tokens.Count > 1024) throw new FormatException("Expression complexity limit exceeded.");
                offset += match.Length;
            }
        }
        private static bool Bare(string token)=>!token.StartsWith('"')&&!token.StartsWith('\'')&&!token.StartsWith('[')&&token is not ("(" or ")")&&Precedence(token)<0&&!token.Equals("NOT",StringComparison.OrdinalIgnoreCase);

        public string Parse()
        {
            if (tokens.Count == 0) return "";
            var result = Binary(0);
            if (index != tokens.Count) throw new FormatException($"Unexpected expression token '{tokens[index]}'.");
            return result;
        }

        private string Binary(int minimum)
        {
            var left = Unary();
            while (index < tokens.Count && Precedence(tokens[index]) >= minimum)
            {
                var op = tokens[index++].ToUpperInvariant();
                var right = Binary(Precedence(op) + (op == "^" ? 0 : 1));
                left = Apply(op, left, right);
            }
            return left;
        }

        private string Unary()
        {
            if (index >= tokens.Count) throw new FormatException("Expression operand is missing.");
            var token = tokens[index++];
            if (token == "(")
            {
                var result = Binary(0);
                if (index >= tokens.Count || tokens[index++] != ")") throw new FormatException("Missing closing parenthesis.");
                return result;
            }
            if (token == "-") return Format(-Number(Unary()));
            if (token == "+") return Format(Number(Unary()));
            if (token.Equals("NOT", StringComparison.OrdinalIgnoreCase)) return Truth(Unary()) ? "0" : "1";
            if (token.StartsWith("\"", StringComparison.Ordinal) || token.StartsWith("'", StringComparison.Ordinal))
                return token.Substring(1, token.Length - 2).Replace("\"\"", "\"");
            if (token.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return "1";
            if (token.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return "0";
            if(token.StartsWith("&H",StringComparison.OrdinalIgnoreCase))return Convert.ToInt64(token[2..],16).ToString(CultureInfo.InvariantCulture);
            if(token.StartsWith('['))token=token[1..^1];
            return variables.TryGetValue(token, out var value) ? value : token;
        }

        private static int Precedence(string op) => op.ToUpperInvariant() switch
        {
            "OR" => 0, "XOR" => 1,"AND" => 2, "=" or "<>" or "<" or ">" or "<=" or ">=" => 3,
            "&"=>4,"+" or "-" => 5, "*" or "/" or "\\" or "MOD" => 6, "^" => 7, _ => -1
        };

        private static string Apply(string op, string left, string right)
        {
            if (op == "&") return left + right;
            if (op == "AND") return Truth(left) && Truth(right) ? "1" : "0";
            if (op == "OR") return Truth(left) || Truth(right) ? "1" : "0";
            if (op == "XOR") return Truth(left) ^ Truth(right) ? "1" : "0";
            if (op == "=" || op == "<>" || op == "<" || op == ">" || op == "<=" || op == ">=")
            {
                var comparison = double.TryParse(left, NumberStyles.Float, CultureInfo.InvariantCulture, out var a)
                    && double.TryParse(right, NumberStyles.Float, CultureInfo.InvariantCulture, out var b)
                    ? a.CompareTo(b) : string.Compare(left, right, StringComparison.Ordinal);
                return (op switch { "=" => comparison == 0, "<>" => comparison != 0, "<" => comparison < 0,
                    ">" => comparison > 0, "<=" => comparison <= 0, _ => comparison >= 0 }) ? "1" : "0";
            }
            var x = Number(left);
            var y = Number(right);
            if ((op == "/" || op == "\\" || op == "MOD") && y == 0) throw new DivideByZeroException();
            return Format(op switch
            {
                "+" => x + y, "-" => x - y, "*" => x * y, "/" => x / y,
                "\\" => Math.Truncate(x / y), "MOD" => x % y, "^" => Math.Pow(x, y),
                _ => throw new FormatException($"Unknown operator {op}.")
            });
        }
    }
}
