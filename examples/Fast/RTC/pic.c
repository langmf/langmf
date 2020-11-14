
Gray_Image(ptr, Width, Height, X1, Y1, X2, Y2) {	
	int x,y,p,b;

	for (y=Y1; y<Y2; y++) {
		p = ptr + 4 * ((Height - y) * Width + (X1-1));
		for (x=X1; x<X2; x++) {
			b = *(char *)p++;
			*(char *)p++ = b;
			*(char *)p++ = b;
			*(char *)p++ = 0;
		}
	}
}

Red_Image(ptr, Width, Height, X1, Y1, X2, Y2) {	
	int x,y,p,b;

	for (y=Y1; y<Y2; y++) {
		p = ptr + 4 * ((Height - y) * Width + (X1-1));
		for (x=X1; x<X2; x++) {
			*(int *)p = *(int *)p & 0xFF0000;
			p = p + 4;
		}
	}
}
