<script>
	/** @type {import('./$types').PageData} */
	export let data;
	const { class_room_data } = data;
	import ComboBox from '../../lib/components/ComboBox.svelte';
	function generateMonthData(numMonths) {
		const currentDate = new Date();
		let currentMonth = currentDate.getMonth() + 1; // Month is zero-based, so add 1
		let currentYear = currentDate.getFullYear();
		const monthsData = [];

		// Generate data for the current month
		monthsData.push({
			text: `${10}/${currentMonth}/${currentYear} - ${9}/${currentMonth + 1}/${currentYear}`,
			value: `${10}/${currentMonth}/${currentYear}`
		});

		// Generate data for previous months, maximum 3
		for (let i = 1; i <= 3; i++) {
			const previousMonth = currentMonth === 1 ? 12 : currentMonth - 1;
			const previousYear = currentMonth === 1 ? currentYear - 1 : currentYear;

			monthsData.unshift({
				text: `${10}/${previousMonth}/${previousYear} - ${9}/${currentMonth}/${currentYear}`,
				value: `${10}/${previousMonth}/${currentYear}`
			});

			// Update currentMonth and currentYear for the next iteration
			currentMonth = previousMonth;
			currentYear = previousYear;
		}
		currentMonth = currentDate.getMonth() + 2; // Month is zero-based, so add 1
		currentYear = currentDate.getFullYear();
		// Generate data for next months
		for (let i = 0; i < numMonths - 4; i++) {
			const nextMonth = currentMonth === 12 ? 1 : currentMonth + 1;
			const nextYear = currentMonth === 12 ? currentYear + 1 : currentYear;

			monthsData.push({
				text: `${10}/${currentMonth}/${currentYear} - ${9}/${nextMonth}/${nextYear}`,
				value: `${10}/${currentMonth}/${currentYear}`
			});

			// Update currentMonth and currentYear for the next iteration
			currentMonth = nextMonth;
			currentYear = nextYear;
		}

		return monthsData;
	}

	// Usage example:
	const fromDateData = generateMonthData(7);

	export let form;
</script>

{#if form?.error}
	<h2>Failed to upload, wrong format!</h2>
{/if}

{#if form?.success}
	<h2>File Uploaded!</h2>
{/if}
<h1>Upload File</h1>
<div class="form">
	<form method="post" action="?/upload" enctype="multipart/form-data">
		<label for="file-upload">Chọn file:</label>
		<input type="file" id="file-upload" name="file-upload" accept="*/*" required />
		<label for="worksheet-name">Chọn tên của worksheet chứa danh sách học sinh:</label>
		<input type="text" id="worksheet-name" name="worksheet_name" placeholder="Sheet1" required />
		<div class="stack">
			<ComboBox
				label="Chọn lớp học"
				name="class_room_id"
				placeholder="Type to search..."
				options={class_room_data}
			/>
			<ComboBox
				label="Chọn khoảng thời gian"
				name="fromDate"
				placeholder="Type to search..."
				options={fromDateData}
				readonly={false}
			/>
		</div>
		<button type="submit">Upload</button>
	</form>
</div>

<style>
	.stack {
		display: flex;
		flex-direction: column;
		gap: 1.5rem;
	}
</style>
