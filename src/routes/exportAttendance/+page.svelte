<script>
	import ComboBox from '../../lib/components/ComboBox.svelte';
	import DatePicker from '../../lib/components/DatePicker.svelte';
	import { writeFileXLSX } from 'xlsx';
	import ExcelJS from 'exceljs';
	/** @type {import('./$types').PageData} */
	export let data;
	import { enhance } from '$app/forms';
	const { class_room_data } = data;
	const currentDate = new Date();

	// Từ ngày 10 của tháng trước ngày hiện tại
	const date1 = new Date(currentDate);
	date1.setDate(10); // Đặt ngày là ngày 10 của tháng hiện tại
	date1.setMonth(date1.getMonth() - 1); // Trừ một tháng để lấy tháng trước

	// Từ ngày 9 của tháng này
	const date2 = new Date(currentDate);
	date2.setDate(9); // Đặt ngày là ngày 9 của tháng hiện tại

	// Từ ngày 10 của tháng này đến tháng sau
	const date3 = new Date(currentDate);
	date3.setDate(10); // Đặt ngày là ngày 10 của tháng hiện tại
	date3.setMonth(date3.getMonth()); // Cộng một tháng để lấy tháng sau

	// Từ ngày 9 của tháng sau
	const date4 = new Date(currentDate);
	date4.setDate(9); // Đặt ngày là ngày 9 của tháng hiện tại
	date4.setMonth(date4.getMonth() + 1); // Cộng một tháng để lấy tháng sau

	function formatDate(date) {
		const day = date.getDate();
		const month = date.getMonth() + 1;
		const year = date.getFullYear();
		return (day < 10 ? '0' : '') + day + '/' + (month < 10 ? '0' : '') + month + '/' + year;
	}

	let day10LastMonth = formatDate(date1);

	let day9ThisMonth = formatDate(date2);

	let day10ThisMonth = formatDate(date3);

	let day9NextMonth = formatDate(date4);
	// Helper function to convert ExcelJS workbook to base64
	export let form;
</script>

<h1>Tải file</h1>
<div class="form">
	<form
		method="post"
		action="?/download"
		use:enhance={({ formElement, formData, action, cancel, submitter }) => {
			// `formElement` is this `<form>` element
			// `formData` is its `FormData` object that's about to be submitted
			// `action` is the URL to which the form is posted
			// calling `cancel()` will prevent the submission
			// `submitter` is the `HTMLElement` that caused the form to be submitted

			return async ({ result, update }) => {
				let buffer = result.data['workbook'];
				// Create a blob from the base64 string
				const byteCharacters = atob(buffer);
				const byteNumbers = new Array(byteCharacters.length);
				for (let i = 0; i < byteCharacters.length; i++) {
					byteNumbers[i] = byteCharacters.charCodeAt(i);
				}
				const byteArray = new Uint8Array(byteNumbers);
				const blob = new Blob([byteArray], {
					type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,'
				});

				// Create a temporary URL for the blob
				const url = window.URL.createObjectURL(blob);

				// Create a link element to trigger the download
				const a = document.createElement('a');
				a.href = url;
				a.download = 'excel_file.xlsx';
				a.style.display = 'none';

				// Append the link to the DOM and trigger the click event
				document.body.appendChild(a);
				a.click();

				// Clean up by removing the link and revoking the blob URL
				document.body.removeChild(a);
				window.URL.revokeObjectURL(url);
			};
		}}
	>
		<div class="stack">
			<ComboBox
				label="Chọn lớp học"
				name="class_room_id"
				placeholder="Type to search..."
				options={class_room_data}
			/>
		</div>
		<button type="submit">Tải file điểm danh từ {day10LastMonth} đến {day9ThisMonth}</button>
		<button type="submit">Tải file điểm danh từ {day10ThisMonth} đến {day9NextMonth}</button>
	</form>
</div>

<style>
	.stack {
		display: flex;
		flex-direction: column;
		gap: 1.5rem;
	}
</style>
