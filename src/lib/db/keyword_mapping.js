export const attributeOrder = 
['student.grade', 'student.first_name', 'student.last_name', 'student.enroll_date', 'student.dob',
  'father.phone_number', 'mother.phone_number', 'father.zalo', 'student.birth_year', 'student.sex',
  'student.ethnic', 'student.birth_place', 'father.name', 'father.dob', 'father.occupation',
  'mother.name', 'mother.dob', 'mother.occupation', 'student.perm_res_province', 'student.perm_res_district',
  'student.perm_res_commune', 'student.temp_res', 'father.landlord', 'father.roi', 'father.birthplace',
  'father.res_registration'
]


export const attributeMappings = {

  'student': {
    'grade': "LỚP",
    'first_name': 'HỌ VÀ TÊN LÓT',
    'last_name': 'TÊN HỌC SINH',
    'enroll_date': 'NGÀY NHẬP HỌC',
    'dob': 'SINH NGÀY',
    'birth_year': 'NĂM SINH',
    'sex': 'GIỚI TÍNH',
    'ethnic': 'DÂN TỘC',
    'birth_place': 'NƠI SINH HS',
    'temp_res': 'TẠM TRÚ',
    'perm_res_province': 'THƯỜNG TRÚ.TỈNH',
    'perm_res_district': 'THƯỜNG TRÚ.HUYỆN',
    'perm_res_commune': 'THƯỜNG TRÚ.XÃ'
  },
  'parent': {
    'phone_1': 'SỐ ĐIỆN THOẠI.ĐT1',
    'phone_2': 'SỐ ĐIỆN THOẠI.ĐT2',
    'zalo': 'ZALO',
    'father_name': 'THÔNG TIN CHA.HỌ VÀ TÊN',
    'father_dob': 'THÔNG TIN CHA.NĂM SINH',
    'father_occupation': 'THÔNG TIN CHA.NGHỀ NGHIỆP',
    'mother_name': 'THÔNG TIN MẸ.HỌ VÀ TÊN',
    'mother_dob': 'THÔNG TIN MẸ.NĂM SINH',
    'mother_occupation': 'THÔNG TIN MẸ.NGHỀ NGHIỆP',
    'landlord': 'CHỦ NHÀ TRỌ',
    'roi': 'BC PHỔ CẬP',
    'birthplace': 'KHAI SINH',
    'res_registration': 'HỘ KHẨU'
  }
};

