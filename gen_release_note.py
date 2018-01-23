#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Copyright (C) 2018 by wranglerPENG https://github.com/pzhaoyang

import xlwt
from datetime import datetime

# static segment
A1Name = 'Project Summary'
A2Name = 'Project'
A3Name = 'Git/SVN Repository'
A4Name = 'Git/SVN Branch'
A5Name = 'svn/git Revision id'
A6Name = 'SW Version'
A7Name = 'Release Note'
A8Name = u'序号'
B8Name = u'需求'


# dynamic value
PrjValue='P50_YUHO'
RepoAddr='git@tpgithost:Android70/MT6739N_V1.git'
BranchValue='P50_YUHO'
LastCommit='git log -1'
VersionValue='YUHO_O1_V1.0_20171227'


# create a workbook
wb = xlwt.Workbook()
# wb = xlwt.Workbook(encoding = 'ascii')

#add a sheet
wsi = wb.add_sheet('Internal')

styleStaticName = xlwt.easyxf('font: name Times New Roman, height 240, color-index black, bold off; \
                               align: vert centre, horiz center')
# static write
wsi.write(0, 0, A1Name,styleStaticName)
wsi.write(1, 0, A2Name,styleStaticName)
wsi.write(2, 0, A3Name,styleStaticName)
wsi.write(3, 0, A4Name,styleStaticName)
wsi.write(4, 0, A5Name,styleStaticName)
wsi.write(5, 0, A6Name,styleStaticName)
wsi.write(6, 0, A7Name,styleStaticName)
wsi.write(7, 0, A8Name,styleStaticName)
wsi.write(7, 1, B8Name,styleStaticName)
wsi.col(0).width = 3333*2
wsi.col(1).width = 3333*5
for i in range(0, 7):
    wsi.row(i).height = 20*20



# daynamic write
wsi.write(1, 1, PrjValue)
wsi.write(2, 1, RepoAddr)
wsi.write(3, 1, BranchValue)
wsi.write(4, 1, LastCommit)
wsi.write(5, 1, VersionValue)


## save boot to file
wb.save('ReleaseNote.xls')