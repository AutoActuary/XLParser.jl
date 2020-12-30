using XLParser
using Test


@testset "Sanitize Excel strings" begin

    @test XLParser._sanitize_excel_formula_strings("\"hello\"") == "raw\"hello\""

    @test(
        XLParser._sanitize_excel_formula_strings(
            "1 & \"hello\" & \"world\" & 2"
        ) == "1 & raw\"hello\" & raw\"world\" & 2"
    )

    # use raw string examples without any escaping to improve readibility
    for line in readlines((@__DIR__)*"/xl_formula_string_examples.txt")
        strip(line) == "" && continue
        i, j = strip.(split(line, "=>"))

        @test XLParser._sanitize_excel_formula_strings(i) == j
    end

end


@testset "Excel to Julia expression" begin

    # dummy functions and variables to use in the examples
    q(args...) = 1
    vvv(args...) = 2
    z = 1:20

    @test xl"False = True" == false

    @test xlexpr"INDEX(z,5):INDEX(z,10)" == :(z[5:10])

    @test xl"INDEX(z,5):INDEX(z,10)" == [5,6,7,8,9,10]

    @test xlexpr"a & b & c & d & e" == :(string(a, b, c, d, e))

    # for some reason the expressions isn't equal although they are exactly the
    # same and even have the same string representation:
    @test(
        string(xlexpr""" Q(1,2,,4) <> vvv(1e-8, "a""b""c" & "!" & INDEX(z,1)) & 2 """)==
        string(:(q(1, 2, nothing, 4) != string(vvv(1.0e-8, string(raw"a\"b\"c", raw"!", z[1])), 2)))
    )

    @test xl"Q() & true" == "1true"

    @test xl"Q() & true <> Q() & true" == false

    @test xl"Q() & true = Q() & true" == true

    @test xl""" "L" & 10*3+3 & "t" """ == "L33t"
end
