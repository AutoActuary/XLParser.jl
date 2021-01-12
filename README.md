# XLParser

This package provide a minimal parser to translate an Excel formula string to a Julia expression.

### Exports
- `@xlexpr_str` string macro to convert Excel formula to Julia expression
- `xlexpr` function to convert Excel formula to Julia expression
- `@xl_str` string macro to run an Excel formula directly 

### Basic Examples
```
julia> using XLParse

julia> xlexpr"foo(1,2,3)" # translate macro
:(foo(1, 2, 3))

julia> xlexpr" 1 & 2 & 3" # translate macro
:(string(1, 2, 3))

julia> xlexpr(" 1 & 2 & 3") # translate function
:(string(1, 2, 3))

julia> xl" 1 & 2 & 3" # translate and eval macro
"123"
```

### Installing this package

Currently, the package is open source but not in the Julia registries 
as it is mostly used internally. To use thus package, install it via 
the github link by opening the Julia REPL and pressing `]` to enter pkg mode:
```
julia> ] 
(@v1.6) pkg> add https://github.com/AutoActuary/XLParser.jl 
```

### Development environment
If you are an Auto Actuary developer who needs to develop on this package,
we recommend installing Julia through [Juliawin](https://github.com/heetbeet/juliawin).

To add this package to your Julia installation, enter pkg mode in Julia and install a dev environment:
```
julia> ] 
(@v1.6) pkg> dev git@github.com:AutoActuary/XLParser.jl.git
```

This will clone and install the package on your current Julia. You can
now find the location of this local repo using:
```
julia> using XLParser 

julia> println(pathof(XLParser))
C:\Users\simon\Juliawin-1.6\userdata\.julia\dev\XLParser\src\XLParser.jl 
```
