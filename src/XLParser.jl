module XLParser
    export @xlexpr_str
    export @xl_str
    export xlexpr

    using Tokenize
    using Espresso

    # we bait and switch string concatenation (ampersand in xl) with the pipe
    # character, since it is placed on the appropriate precedence level in Julia:
    # https://github.com/JuliaLang/julia/blob/master/src/julia-parser.scm
    dummy_ampersand = :(|>)

    function _sanitize_excel_formula_strings(str)
        chars = collect(str)
        str_builder = Array{Union{String, Char}, 1}()

        i = 0
        inquote = false
        while (i+=1) <= length(chars)

            # Mark the beginning of a string
            if !inquote && chars[i] == '"'
                inquote = true
                push!(str_builder, "raw\"")

            # escaping backslashes if they are before a quote (double them up)
            elseif inquote && chars[i] == '\\'
                cnt = 1
                while (i+=1) <= length(chars) && chars[i] == '\\'
                    cnt+=1
                end

                if chars[i] == '"'
                    push!(str_builder, "\\"^(cnt*2))
                else
                    push!(str_builder, "\\"^(cnt))
                end

                # reset our peek-forward
                i -= 1

            # escaping of a quote
            elseif inquote && (i+1) <= length(chars) && chars[i] == '"' && chars[i+1] == '"'
                push!(str_builder, "\\\"")
                i+=1

            # closing the string
            elseif inquote && chars[i] == '"'
                push!(str_builder, '"')
                inquote = false

            # normal passthrough
            else
                push!(str_builder, chars[i])
            end
        end

        return join(str_builder, "")
    end


    function _sanitize_excel_formula_tokens(str)
        # ensure excel strings are converted to Julia syntax
        str = _sanitize_excel_formula_strings(str)
        str = replace(str, "\n" => " ")

        tokens = collect(tokenize(str))

        i = 0
        while (i+=1) <= length(tokens)

            # convert names to lowercase
            if tokens[i].kind == Tokenize.Tokens.IDENTIFIER && tokens[i].val != lowercase(tokens[i].val)
                tokens[i] = first(tokenize(lowercase(tokens[i].val)))
                i-=1

            # convert if KEYWORD token to xl_if function token
            elseif tokens[i].kind == Tokenize.Tokens.IF
                tokens[i] = first(tokenize("xl_if"))


            # convert <> to !=
            elseif (i != length(tokens) &&
                tokens[i].kind == Tokenize.Tokens.LESS &&
                tokens[i+1].kind == Tokenize.Tokens.GREATER)

                tokens[i] = first(tokenize("!="))
                tokens[i+1] = first(tokenize(""))

            # convert = to ==
            elseif (tokens[i].kind == Tokenize.Tokens.EQ)
                tokens[i] = first(tokenize("=="))

            # add whitespace between repeating minus signs to avoid syntax error
            elseif (tokens[i].token_error == Tokenize.Tokens.INVALID_OPERATOR &&
                        tokens[i].val == "--")
                tokens[i] = first(tokenize("-"))
                insert!(tokens, i+1, first(tokenize(" ")))
                insert!(tokens, i+1, first(tokenize("-")))
                insert!(tokens, i+1, first(tokenize(" ")))

            # fill empty commas: FOO(,,) -> FOO(nothing,nothing,nothing)
            elseif (tokens[i].kind == Tokenize.Tokens.COMMA)
                # test if left of comma is bracket
                j = i-1
                if j >= 1 && tokens[j].kind == Tokenize.Tokens.WHITESPACE
                    j -= 1
                end

                if j >= 1 && tokens[j].kind == Tokenize.Tokens.LPAREN
                    insert!(tokens, j+1, first(tokenize("nothing")))
                    i += 1
                end

                # test if right of comma is another comma or )
                j = i+1
                if j <= length(tokens) && tokens[j].kind == Tokenize.Tokens.WHITESPACE
                    j += 1
                end

                if j <= length(tokens) && (tokens[j].kind == Tokenize.Tokens.COMMA ||
                                           tokens[j].kind == Tokenize.Tokens.RPAREN)
                    insert!(tokens, j, first(tokenize("nothing")))
                    i += 1
                end

            # Replace ampersand with a lower precedence infix operator
            elseif  tokens[i].kind == Tokenize.Tokens.AND
                tokens[i] = first(tokenize(string(dummy_ampersand)))
            end
        end

        return join(untokenize.(tokens), "")
    end


    function xlexpr(str)
        # turn string into a valid (but logically wrong) Julia expression
        parsable_string = _sanitize_excel_formula_tokens(str)
        xl = Meta.parse(parsable_string)

        # now correct this logically wrong expression into a correct one

        # index syntax, eg: a[1:5]
        xl = rewrite_all(xl, :(index(_a,_b):index(_a,_d)), :(_a[_b:_d]))
        xl = rewrite_all(xl, :(index(_a,_b)), :(_a[_b]))

        #TODO: test expr contains index(_a,_b):index(_c,_d) and throw an error if so
        xl = rewrite_all(xl, :(index(_a,_b):index(_a,_d)), :(_a[_b:_d]))

        # replace & infix with string function
        xl = rewrite_all(xl, dummy_ampersand, :(string))
        xl = rewrite_all(xl, :(string(string(_a...), _b)),
                             :(string(_a..., _b)))

        # replace and(1,2,3) with 1 && (2 && 3)
        while xl != (xl_new = rewrite_all(xl, :(and(_a, _b...)),
                                              :(_a && and(_b...))))
            xl = xl_new
        end
        xl = rewrite_all(xl, :(and(_a)), :(_a))


        # replace or(1,2,3) with 1 || (2 || 3)
        while xl != (xl_new = rewrite_all(xl, :(or(_a, _b...)),
                                              :(_a || or(_b...))))
            xl = xl_new
        end
        xl = rewrite_all(xl, :(or(_a)), :(_a))

        # replace not(1) with !(1)
        xl = rewrite_all(xl, :(not(_a)), :(!(_a)))

        # replace excel if functions with julia if statements
        xl = rewrite_all(xl, :(xl_if(_a,_b,_c)), :((_a ? _b : _c)))
        xl = rewrite_all(xl, :(xl_if(_a,_b)), :((_a ? _b : false)))

        return xl
    end


    macro xlexpr_str(str)
        # quot allows macro to not execute but to return the expression
        Meta.quot(xlexpr(str))
    end


    macro xl_str(str)
        # esc allows macro to execute in caller's scope
        esc(xlexpr(str))
    end

end
