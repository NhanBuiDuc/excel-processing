<script>
	import { readFile, utils, writeFileXLSX } from 'xlsx';
	// /** @type {import('./$types').PageData} */
	// export let data;
	let pres = [];
	let baseUrl = 'http://127.0.0.1:5173/';
	let studentListFileName = 'template_danhsachhocsinh.xlsx';
	let attendanceFileName = 'template_diemdanh.xlsx';
	let studentListUrl = '';
	let attendanceUrl = '';
	async function exportStudentListTemplate() {
		studentListUrl = baseUrl + studentListFileName;
	}
	function exportAttendanceTemplate() {
		attendanceUrl = baseUrl + attendanceFileName;
	}
</script>

<div>
	<button on:click={exportStudentListTemplate}>
		<a href={studentListUrl}> Tải mẫu danh sách học sinh </a>
	</button>
	<button on:click={exportAttendanceTemplate}>
		<a href={attendanceUrl}> Tải file điểm danh mẫu</a>
	</button>
</div>
