function mouseDownDragHandler(element, event) {
	element.style.cursor = 'grabbing';
	element.style.userSelect = 'none';
	element.pos = {
		left: element.scrollLeft,
		top: element.scrollTop,
		x: event.clientX,
		y: event.clientY
	}

	element.mouseMoveListener = evt => mouseMoveDragHandler(element, evt);
	element.mouseUpListener = evt => mouseUpDragHandler(element);
	element.addEventListener("mousemove", element.mouseMoveListener);
	element.addEventListener("mouseup", element.mouseUpListener);
	element.addEventListener("mouseout", element.mouseUpListener);
}

function mouseMoveDragHandler(element, event) {
	if (element.pos) {
		const dx = event.clientX - element.pos.x;
		const dy = event.clientY - element.pos.y;
		element.scrollTop = element.pos.top - dy;
		element.scrollLeft = element.pos.left - dx;
	}
}

function mouseUpDragHandler(element) {
	element.pos = null;
	element.style.cursor = 'grab';
	element.style.removeProperty('user-select');

	if (element.mouseMoveListener) {
		element.removeEventListener("mousemove", element.mouseMoveListener);
	}

	if (element.mouseUpListener) {
		element.removeEventListener("mouseup", element.mouseUpListener);
		element.removeEventListener("mouseout", element.mouseUpListener);
	}
}

