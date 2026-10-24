#!/bin/bash
cd /tmp/claude-0/-home-user-FARP/2cd27235-ceb3-518d-bee2-e3b544119c0a/scratchpad
export NODE_PATH=$(npm root -g)
enc(){ ffmpeg -y -loglevel error -framerate 12 -i vid/s$1/f%05d.jpg -vf "minterpolate=fps=24:mi_mode=mci:mc_mode=aobmc:me_mode=bidir:vsbmc=1:scd=fdiff:scd_threshold=12" -c:v libx264 -pix_fmt yuv420p -crf 18 -movflags +faststart vid/s$1.mp4; }
node render.js 4 6 21 12 8783 && enc 04
node render.js 10 0 36 12 8783 && enc 10
node render.js 5 0 32 12 8783 && enc 05
node render.js 11 0 32 12 8783 && enc 11
node render.js 12 0 18 12 8783 && enc 12
echo ALLDONE_C
