#!/bin/tcsh

# Set up python and cplex environment

conda activate /usr/local/usrapps/infews/CAPOW_env
module load gurobi
source /usr/local/apps/gurobi/gurobi810/linux64/bin/gurobi.sh

# Submit multiple jobs at once
@ totalFols = 69

set folNameBase = NEISO_110_0_

@ folNum = 0

while ($folNum < $totalFols)

	set dirName = ${folNameBase}${folNum}
   	cd $dirName

    	# Submit LSF job for the directory $dirName
   	bsub -n 8 -R "span[hosts=1]" -W 5000 -x -o out.%J -e err.%J "python NEISO_simulation.py"

	# Go back to upper level directory
    	cd ..

    	@ folNum = $folNum + 1
end

conda deactivate

	
