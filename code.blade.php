<?php

/**************************** This code is for generating excel **************************/

// <!-- {{-- this is from controller --}} -->
class ReportsController extends Controller
    {
        public function getReports(Request $request)
        {
            $user_id        = Session::get('user');
            $user           = User::where('partner_id',$user_id)->first();
            $password       = DB::table('users')->where('partner_id',$user_id)->first();
            $tabbing        = ['tab' => 'Reports', 'subtab' => ''];

            return view('dgm.report.index');
        }

        public function getDownloadExcelGroupSession(Request $request)
        {
            $group_members = DgmGroupMember::where('dgm_group_members.status', 'Active')->select('dgm_group_members.*')->orderBy('name', 'asc')->orderBy('action_status', 'asc')->get();
            
            $date_from = $request->input('date_from');
            $date_to = $request->input('date_to');

            ob_end_clean();
            ob_start();

            Excel::create('Group Sessions Reports'. date('M d, Y', strtotime($date_from)) . ' to ' . date('M d, Y', strtotime($date_to)), function($excel) use($group_members, $date_from, $date_to) {
                $excel->sheet('Groups Report', function($sheet) use($group_members, $date_from, $date_to) {

                    $sheet->mergeCells('A1:N1');
                    $sheet->getStyle('A1:N1')->getAlignment()->applyFromArray(array('horizontal' => 'center'));
                    $sheet->row(1, array(
                        'Group Sessions Report (' . date('M d, Y', strtotime($date_from)) . ' to ' . date('M d, Y', strtotime($date_to)) . ')'
                    ));

                    $sheet->row(1, function($row) {
                        $row->setFontWeight('bold');
                    });

                    $header = ['Group ID', 'Life Stage', 'Dgroup Name', 'Active Members', 'Maximum Members', 'Accepting Seekers', 'Language Spoken', 'Male Dleader', 'Female Dleader', 'Schedule', 'Time', 'Meeting Area', 'Date Created', 'No. of Sessions'];

                    $sheet->row(3,array_merge($header));

                    $sheet->row(3, function($row) {
                        $row->setFontWeight('bold');
                    });

                    $row_line = 4;

                    foreach($group_members as $group_member){
                        $group_member_session = $group_member->get_sessions($group_member->id, $date_from, $date_to);

                        if($group_member->get_life_stage)
                        {
                            $life_stage_name = $group_member->get_life_stage->name;
                        }

                        $active_member = $group_member->total_group_member_count($group_member->id);

                        $lang_spoken = [];
                        
                        if($group_member->get_language_spokens){
                            foreach($group_member->get_language_spokens as $key => $group_member_language)
                            {
                                if(count($group_member->get_language_spokens) == 1)
                                {
                                    $lang_spoken[] = $group_member_language->get_language->name;
                                }
                                else
                                {
                                    if(count($group_member->get_language_spokens) == ($key+1)){
                                        $lang_spoken[] = $group_member_language->get_language->name;
                                    }
                                    else
                                    {
                                        $lang_spoken[] = $group_member_language->get_language->name;
                                    }
                                }
                            }
                        }

                        
                        $schedule = $group_member->get_schedule($group_member->schedule);

                        $languages = implode(', ', $lang_spoken);

                        
                        $arr = [
                            str_pad($group_member->id,5,'0',STR_PAD_LEFT),
                            $life_stage_name,
                            $group_member->name,
                            $active_member,
                            $group_member->maximum_members,
                            ($group_member->group_type == 'Open' ? 'Yes' : 'No'),
                            $languages,
                            ($group_member->get_male_dleader ? $group_member->get_male_dleader->full_name_second : ''),
                            ($group_member->get_female_dleader ? $group_member->get_female_dleader->full_name_second : ''),
                            $schedule,
                            date("h:i:s A", strtotime($group_member->time)),
                            ($group_member->get_meeting_area ? $group_member->get_meeting_area->name : ''),
                            date("Y-m-d", strtotime($group_member->created_at)),
                            count($group_member_session)
                        ];

                        $lang_spoken = [''];
                        $sheet->row($row_line+1,$arr);

                        $row_line++;
                    }

                });
            })->export('xlsx');
        }
    }
// {{-- end of controller --}}

// this is for the blade
<!-- For Group Session  Report -->
	<div class="modal fade" id="group_session" role="dialog" data-backdrop="false">
		<div class="modal-dialog" style="margin-top:3%;width:30%;">
			<!-- Modal content-->
			<div class="modal-content">
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal">&times;</button>
					<h4 class="modal-title"><b>Sales Report</b></h4>
				</div>
				
				<div class="modal-body">
					<form class="form-horizontal salesreport_form" id="salesreport_form" autocomplete="off" name="salesreport_form" target="_blank" role="form"  method="POST" action="{{ route('dgm.reports.excel')}}" >
					<input type="hidden" name="_token" class="token" value="{{ csrf_token() }}">
					<div class="form-group" style="margin-top:-15px;height:23px;width:107.4%;">
                    	<label class="col-md-12 control-label modalmessage1" style="text-align: left;width:107.4%;margin-left:-13px;margin-top:-6px;font-size: 10pt;"></label>
					</div>
					<div class="form-group">
                        <label class="col-md-4 control-label">Start Date:</label>
                        <div class="col-md-8">
                            <input type="text" class="form-control date_from" name="date_from" id="date_from">
                        </div>
                    </div>
                    <div class="form-group">
                        <label class="col-md-4 control-label">End Date:</label>
                        <div class="col-md-8">
                            <input type="text" class="form-control date_to" name="date_to" id="date_to">
                        </div>
                    </div>
				</form>
				</div>

				<div class="modal-footer">
					<div style="float: right;">
						<button class="btn btn-sm printsalesreport_excel" type="submit" form="salesreport_form" id="printsalesreport_excel" name="printsalesreport_excel" value="printsalesreport_excel" style="margin-top:-7px;background-color:#a6a6a6;font-weight:200;color:#0d0d0d;height:28px;border:1px solid #8c8c8c;"><img src="\images\downloads.png">&nbsp;Excel</button>
					</div>
				</div>
			</div>
		</div>
	</div>
	<!-- end For print Sales Report -->

// end of the blade

// This is for routes
Route::post('reports/group_session_report', 'ReportsController@getDownloadExcelGroupSession')
    ->name('dgm.reports.excel');
// end for the routes

/**************************** end code for generating excel **************************/