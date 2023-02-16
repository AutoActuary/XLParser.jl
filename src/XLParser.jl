module XLParser
export @xlexpr_str
export @xl_str
export xlexpr

using Tokenize
using Espresso
using MacroTools: MacroTools, @capture, postwalk

# we bait and switch string concatenation (ampersand in xl) with the pipe
# character, since it is placed on the appropriate precedence level in Julia:
# https://github.com/JuliaLang/julia/blob/master/src/julia-parser.scm
dummy_ampersand = "|>"

# same with percentage and ::ₓₓ℅ₓₓᵖᵉʳᶜᵉⁿᵗ
dummy_percent = "::ːːˑ℅ˑᵖᵉʳᶜᵉⁿᵗˑːː"

function _pretransform_excel_formula_string(str)
    chars = collect(str)
    str_builder = Array{Union{String,Char},1}()

    i = 0
    inquote = false
    while (i += 1) <= length(chars)

        # Mark the beginning of a string
        if !inquote && chars[i] == '"'
            inquote = true
            push!(str_builder, "raw\"")

            # escaping backslashes if they are before a quote (double them up)
        elseif inquote && chars[i] == '\\'
            cnt = 1
            while (i += 1) <= length(chars) && chars[i] == '\\'
                cnt += 1
            end

            if chars[i] == '"'
                push!(str_builder, "\\"^(cnt * 2))
            else
                push!(str_builder, "\\"^(cnt))
            end

            # reset our peek-forward
            i -= 1

            # escaping of a quote
        elseif inquote && (i + 1) <= length(chars) && chars[i] == '"' && chars[i + 1] == '"'
            push!(str_builder, "\\\"")
            i += 1

            # closing the string
        elseif inquote && chars[i] == '"'
            push!(str_builder, '"')
            inquote = false

            # normal passthrough
        else
            push!(str_builder, chars[i])
        end

        # bait and switch ampersand and percentage operators
        if !inquote && i <= length(chars)
            if str_builder[end] == '%'
                str_builder[end] = dummy_percent
            elseif str_builder[end] == '&'
                str_builder[end] = dummy_ampersand
            end
        end
    end

    return join(str_builder, "")
end

function _transform_excel_formula_string(str)
    # ensure excel strings are converted to Julia syntax
    str = _pretransform_excel_formula_string(str)
    str = replace(str, "\n" => " ")

    tokens = collect(tokenize(str))

    i = 0
    while (i += 1) <= length(tokens)

        # convert names to lowercase
        if tokens[i].kind == Tokenize.Tokens.IDENTIFIER &&
            tokens[i].val != lowercase(tokens[i].val)
            tokens[i] = first(tokenize(lowercase(tokens[i].val)))
            i -= 1

            # convert if KEYWORD token to xl_if function token
        elseif tokens[i].kind == Tokenize.Tokens.IF
            tokens[i] = first(tokenize("xl_if"))

            # convert <> to !=
        elseif (
            i != length(tokens) &&
            tokens[i].kind == Tokenize.Tokens.LESS &&
            tokens[i + 1].kind == Tokenize.Tokens.GREATER
        )
            tokens[i] = first(tokenize("!="))
            tokens[i + 1] = first(tokenize(""))

            # convert = to ==
        elseif (tokens[i].kind == Tokenize.Tokens.EQ)
            tokens[i] = first(tokenize("=="))

            # add whitespace between repeating minus signs to avoid syntax error
        elseif (
            tokens[i].token_error == Tokenize.Tokens.INVALID_OPERATOR &&
            tokens[i].val == "--"
        )
            tokens[i] = first(tokenize("-"))
            insert!(tokens, i + 1, first(tokenize(" ")))
            insert!(tokens, i + 1, first(tokenize("-")))
            insert!(tokens, i + 1, first(tokenize(" ")))

            # fill empty commas: FOO(,,) -> FOO(nothing,nothing,nothing)
        elseif (tokens[i].kind == Tokenize.Tokens.COMMA)
            # test if left of comma is bracket
            j = i - 1
            if j >= 1 && tokens[j].kind == Tokenize.Tokens.WHITESPACE
                j -= 1
            end

            if j >= 1 && tokens[j].kind == Tokenize.Tokens.LPAREN
                insert!(tokens, j + 1, first(tokenize("nothing")))
                i += 1
            end

            # test if right of comma is another comma or )
            j = i + 1
            if j <= length(tokens) && tokens[j].kind == Tokenize.Tokens.WHITESPACE
                j += 1
            end

            if j <= length(tokens) && (
                tokens[j].kind == Tokenize.Tokens.COMMA ||
                tokens[j].kind == Tokenize.Tokens.RPAREN
            )
                insert!(tokens, j, first(tokenize("nothing")))
                i += 1
            end
        end
    end

    return join(untokenize.(tokens), "")
end

function xlexpr(str)
    # turn string into a valid Julia expression (but still containing incorrect dummy operators wrong)
    parsable_string = _transform_excel_formula_string(str)
    xl = Meta.parse(parsable_string)

    # now correct this logically wrong expression into a correct one

    # index syntax, eg: a[1:5]
    postwalk(xl) do x
        if @capture(x, index(a_, i₁_):index(c_, i₂_)) && a != c
            throw(
                ArgumentError(
                    "indexing across different arrays `index(a, i):index(b, j)` is not supported",
                ),
            )
        end
        return x
    end

    xl = postwalk(xl) do x
        @capture(x, index(var_, i₁_):index(var_, i₂_)) &&
            return Expr(:ref, var, Expr(:call, :(:), i₁, i₂))
        return x
    end

    xl = postwalk(xl) do x
        @capture(x, index(var_, i_)) && return Expr(:ref, var, i)
        return x
    end

    # replace |> (& replacement) infix with string function
    xl = postwalk(xl) do x
        # format off
        if x isa Expr &&
            x.head == :call &&
            length(x.args) == 3 &&
            strip(string(x.args[1])) == strip(dummy_ampersand)
            return Expr(:call, :string, x.args[2], x.args[3])
        end
        return x
        # format on
    end

    xl = postwalk(xl) do x
        @capture(x, string(string(a__), b_)) && return Expr(:call, :string, a..., b)
        return x
    end

    xl = MacroTools.prewalk(xl) do x
        if @capture(x, and(a__))
            length(a) == 0 && return x
            length(a) == 1 && return a[1]
            return Expr(:call, :(&&), a[1], Expr(:call, :and, a[2:end]...))
        end
        return x
    end

    # replace or(1,2,3) with 1 || (2 || 3)
    xl = MacroTools.prewalk(xl) do x
        if @capture(x, or(a__))
            length(a) == 0 && return x
            length(a) == 1 && return a[1]
            return Expr(:call, :(||), a[1], Expr(:call, :or, a[2:end]...))
        end
        return x
    end

    # replace not(1) with !(1)
    xl = MacroTools.prewalk(xl) do x
        if @capture(x, not(a_))
            return Expr(:call, :!, a)
        end
        return x
    end

    xl = postwalk(xl) do x
        if x isa Expr && x.head == :call && x.args[1] == :xl_if
            return Expr(:if, [x.args[2:end]..., false, false, false][1:3]...)
        end
        return x
    end

    # Convert "::ːːˑ℅ˑᵖᵉʳᶜᵉⁿᵗˑːː" to "/100"
    xl = postwalk(xl) do x
        if x isa Expr &&
            x.head == :(::) &&
            string(x.args[2]) == replace(dummy_percent, ":" => "")
            return Expr(:call, :/, x.args[1], 100)
        end
        return x
    end

    return xl
end

macro xlexpr_str(str)
    # quot allows macro to not execute but to return the expression
    return Meta.quot(xlexpr(str))
end

macro xl_str(str)
    # esc allows macro to execute in caller's scope
    return esc(xlexpr(str))
end

end
