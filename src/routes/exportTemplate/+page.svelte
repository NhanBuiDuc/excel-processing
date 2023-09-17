<script>
	let baseUrl = 'http://127.0.0.1:5173/';
	let fileName = '/template.xlsx';
	let url = '';

	function exportAttendanceTemplate() {
		url = baseUrl + fileName;
	}
</script>

<div>
	<button on:click={exportAttendanceTemplate}>
		<a href={url}> Tải file điểm danh mẫu</a>
	</button>
</div>
