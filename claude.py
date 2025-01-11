def Insert_Pictures(prs, path, slideIndex, shName, columns, maxPicCnt, whiteRatio, 
                        temp=False, rot=False, startIndex=1, except_list=None):
    slide = prs.slides[slideIndex]
    files = GetFilePath(directoryPath=path, max_files=maxPicCnt, startIndex=startIndex, except_list=except_list)
    
    # Convert image file format
    tempfiles = []
    
    if temp:
        if not os.path.exists(path+r'\temp'):
            os.makedirs(path+r'\temp')
        
        for i in range(len(files)):
            originFile = r'{0}\{1}'.format(path, files[i])
            tempFile = r'{0}\temp\temp_{1}'.format(path, files[i])
            img = Image.open(originFile)
            img.save(tempFile)
            tempfiles.append(tempFile)
    else:
        for i in range(len(files)):
            originFile = r'{0}\{1}'.format(path, files[i])
            tempfiles.append(originFile)
    
    # Insert pictures
    for shape in slide.shapes:
        if shape.name == shName:
            tableleft = shape.left
            tabletop = shape.top
            tableWidth = shape.width
            tableHeight = shape.height
            
            rows = maxPicCnt / columns
            picWidth = (tableWidth * (1-whiteRatio)) / columns            
            picHeight = (tableHeight * (1-whiteRatio)) / rows
            
            rowGap = (tableHeight - (picHeight*rows)) / (rows * 2)
            colGap = (tableWidth - (picWidth*columns)) / (columns * 2)
            
            rowGap2 = (tableHeight - (picWidth*rows)) / rows
            
            startleft = tableleft + (colGap if columns > 1 else colGap * 2 * (-1))
            starttop = tabletop + (rowGap if columns > 1 else rowGap2 / 2)
                
            for i in range(len(tempfiles)):
                row = i % columns
                col = int(i / columns)
                
                if rot:
                    # Adjust position for rotated images
                    width_diff = (picWidth - picHeight) / 2
                    left = startleft + (((colGap*2) + picWidth) * row) + width_diff
                    top = starttop + (((rowGap*2) + picHeight) * col) - width_diff
                    pic = slide.shapes.add_picture(tempfiles[i], left=left, top=top, width=picHeight, height=picWidth)
                    pic.rotation = 90
                else:
                    left = startleft + (((colGap*2) + picWidth) * row)
                    top = starttop + (((rowGap*2) + picHeight) * col)
                    pic = slide.shapes.add_picture(tempfiles[i], left=left, top=top, width=picWidth, height=picHeight)