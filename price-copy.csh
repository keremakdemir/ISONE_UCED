# Copy shadow prices at once

mkdir shadow_prices_110_798

@ totalFols = 69

set folNameBase = NEISO_110_798_

@ folNum = 0

while ($folNum < $totalFols)

	set dirName = ${folNameBase}${folNum}
   	cd $dirName
	cp shadow_price.csv /share/infews/kakdemi/shadow_prices_110_798/shadow_price_${dirName}.csv
    	cd ..
    	@ folNum = $folNum + 1
end


	
