@bp.route("/upload/<specialty>", methods=['GET', 'POST'])
@csrf.exempt
def upload_file(specialty):
	if request.method == 'POST':
		def generate():
			spec = Specialty.query.filter_by(name=specialty).first_or_404()
			f = request.files['file']
			f = pd.read_excel(f, engine='openpyxl', sheet_name=specialty, header=0, usecols=[0,1,2])
			for index, row in f.iterrows():
				state = row[0]
				name = row[1]
				dates = ast.literal_eval(row[2])
				d = []
				for date in dates:
					d.append(dt.datetime.strptime(date, '%m/%d/%Y'))
				program = Program(name=name, state=state, specialty=spec)
				interview = Interview(interviewer=program,interviewee=current_user)
				dates = list(map(lambda x: Interview_Date(date=x, interviewer=program,interviewee=current_user, invite=interview,full=False), d))
				interview.dates = dates
				db.session.add(interview)
				db.session.commit()
				yield(str(index))
		return Response(stream_with_context(generate()))
	return render_template('upload.html')