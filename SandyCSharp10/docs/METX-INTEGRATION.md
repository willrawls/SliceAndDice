# MetX.Standard.Strings integration

Sandy.Core references the existing binary builds of William's MetX.Standard.Strings library. The xlg checkout is read-only to this project; Sandy does not rebuild, modify or copy source into it.

## Framework selection

| Sandy.Core target | Referenced local file |
| --- | --- |
| net10.0 | `$(MetXStringsRoot)/net10.0/MetX.Standard.Strings.dll` |
| netstandard2.0 | `$(MetXStringsRoot)/netstandard2.0/MetX.Standard.Strings.dll` |
| netstandard2.1 | `$(MetXStringsRoot)/netstandard2.1/MetX.Standard.Strings.dll` |

The default root resolves to the supplied location:

```text
C:\Users\willr\OneDrive\Data\code\xlg\MetX\MetX.Standard.Strings\bin\Debug
```

The default is expressed relative to the neighboring repositories in `Directory.Build.props`, not hardcoded to a Windows username. The reference is copy-local: consumers and the VSIX carry the required runtime DLL. Their users do not need a checkout of xlg. Building Sandy does require the local binaries.

Override the location when the repositories move or when using another build configuration of MetX:

```powershell
& '.\SandyCSharp10\Verify.ps1' -MetXStringsRoot 'D:\Libraries\MetX.Standard.Strings\bin\Release'

dotnet build '.\SandyCSharp10\SandyCSharp10.slnx' -c Release '-p:MetXStringsRoot=D:\Libraries\MetX.Standard.Strings\bin\Release'
```

The supplied directory must contain all three framework subdirectories for a full solution verification. Missing files produce an actionable build error. No older public NuGet package is silently substituted, and no binary is vendored into the repository.

These are live references to Debug outputs: rebuilding MetX can change the input binaries. Verify again after updating them. For a reproducible release, use an immutable build directory through the override or publish a versioned multi-target package.

## Integration points and semantic boundaries

- Inline `Token` and its aliases call MetX `Tokenizer.TokenAt` directly without allocating a split array.
- Inline `TokenCount` and its aliases call MetX `Tokenizer.TokenCount`.
- Before, After, Except, Extract and FindToken use MetX `AllTokens`, retaining adjacent, leading and trailing empty entries.
- Inline Left and Right use MetX `StringExtensions.Left` and `Right`.
- Inline Mid uses the three-argument MetX `Mid` overload, translating Sandy's one-based position to the library's zero-based offset and bounding the length.

Sandy explicitly requests ordinal, case-sensitive delimiter matching where the MetX API accepts a comparison. AllTokens uses the same case-sensitive splitting. The library's default ignore-case behavior is therefore not inherited accidentally. Sandy still rejects empty delimiters and treats empty input as zero tokens, adapting MetX AllTokens' single-empty-entry result.

The inspected MetX TokensAfter/TokensBefore implementations have different empty-token/comparison edge behavior. Sandy uses AllTokens for these composite operations to preserve its established results. No changes were made to MetX itself.

## Verification

`Verify.ps1` runs the suite separately against the actual Core and MetX assemblies for .NET 10, Standard 2.0 and Standard 2.1. Metadata assertions prevent a Standard test run from silently loading the .NET 10 versions. Boundary cases include missing/empty tokens, repeated and case-sensitive multicharacter delimiters, one-based Mid offsets, oversized substring lengths, extraction mutation and empty-delimiter errors.

The script checks SHA-256 equality between the local input and copied MetX DLL for each framework. It also checks that the VSIX contains the exact supplied .NET 10 MetX DLL.
