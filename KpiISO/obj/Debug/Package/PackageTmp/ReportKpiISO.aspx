<%@ Page Title="" Language="C#" MasterPageFile="~/Site1.Master" AutoEventWireup="true" CodeBehind="ReportKpiISO.aspx.cs" Inherits="KpiISO.ReportKpiISO" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <script type="text/javascript"> 


        jQuery(document).ready(function () {
            $('#in_month').datepicker({
                format: 'mm-yyyy',
                autoclose: true,
                minViewMode: 1,
                //todayHighlight: true
            });
            //$('.dataTable').DataTable();

            $('#in_lob').select2({
                placeholder: "เลือกสายงาน",
                allowClear: true
            });

            $('#in_proj_type').select2({
                placeholder: "เลือกประเภทโครงการ",
                allowClear: true
            });

            $('#in_proj').select2({
                placeholder: "เลือกสายงานก่อน",
                allowClear: true
            });

            getLob();
            bindProjTyp();

            var data = '{in_lob: null, in_proj_type: null}';
            getProj(data);





            $("#in_lob").change(function () {
                //alert(this.value);

                //แนวราบ
                if (this.value === '9b70169d-07d2-447d-bc5e-6f16049cb589') {
                    changeSelectedLowRise();
                }
                //ภูมิภาค
                if (this.value !== '9b70169d-07d2-447d-bc5e-6f16049cb589' && this.value !== null) {
                    changeSelectedRegion();
                }

            });

            $('#in_proj_type').change(function () {
                var $proj_type = $('#in_proj_type');

                var $lob = $("#in_lob");

                var data;
                if ($proj_type.val() === "NULL") {//$proj_type.val() === null ||
                    data = '{in_lob: "' + $lob.val() + '", in_proj_type: null}';
                }
                else {
                    data = '{in_lob: "' + $lob.val() + '", in_proj_type: "' + $proj_type.val() + '"}';
                }

                getProj(data);

            });



        });

        function getLob() {
            $.ajax({
                url: 'ReportKpiISO.aspx/GetLob',
                type: 'POST',
                //data: data_,
                contentType: 'application/json; charset=utf-8',
                dataType: 'json',
                success: function (response) {
                    if (response) {

                        var json_d = response.d;
                        //console.log(json_d);

                        // convert string to JSON
                        json_d = $.parseJSON(json_d);
                        //console.log(json_d);

                        bindLob(json_d);

                    }//end if
                },
                failure: function (response) {
                    alert(response.d);
                }
            });
        }

        function bindLob(json) {
            var $select = $('#in_lob');
            $select.find('option').remove();
            $.each(json, function (key, value) {
                $select.append(`<option value="${value.lob_id}">${value.lob_name}</option>`);

            });
        }

        function bindProjTyp() {
            var json = { "LOW_RISE": "แนวราบ", "HIGH_RISE": "อาคารสูง", "NULL": "แนวราบ + อาคารสูง" };
            var $select = $('#in_proj_type');
            $select.find('option').remove();
            $.each(json, function (key, value) {
                $select.append(`<option value="${key}">${value}</option>`);
                //console.log(value);
            });
        }

        function changeSelectedLowRise() {
            var $proj_type = $('#in_proj_type');
            $proj_type.val("LOW_RISE").change();
            //console.log($select.val());

            var $lob = $("#in_lob");
            //console.log($lob.val());

            var data;
            if ($proj_type.val() === "null") {
                data = '{in_lob: "' + $lob.val() + '", in_proj_type: null}';
            }
            else {
                data = '{in_lob: "' + $lob.val() + '", in_proj_type: "' + $proj_type.val() + '"}';
            }
            console.log(data);

            getProj(data);
        }

        function changeSelectedRegion() {
            //var $radio = $('#inr_lowrise');
            //$radio.prop('disabled', true); 
            var $proj_type = $('#in_proj_type');

            var $lob = $("#in_lob");

            //var data = '{in_lob: "' + $lob.val() + '", in_proj_type: ' + $proj_type.val() === null ? null : $proj_type.val() + '}';
            var data;
            if ($proj_type.val() === "null") {
                data = '{in_lob: "' + $lob.val() + '", in_proj_type: null}';
            }
            else {
                data = '{in_lob: "' + $lob.val() + '", in_proj_type: "' + $proj_type.val() + '"}';
            }
            //console.log(data);

            getProj(data);
        }

        function getProj(data) {
            $.ajax({
                url: 'ReportKpiISO.aspx/GetProj',
                type: 'POST',
                data: data,
                contentType: 'application/json; charset=utf-8',
                dataType: 'json',
                success: function (response) {
                    if (response) {

                        var json_d = response.d;
                        //console.log(json_d);

                        // convert string to JSON
                        json_d = $.parseJSON(json_d);
                        //console.log(json_d);

                        bindProj(json_d);

                    }//end if
                },
                failure: function (response) {
                    alert(response.d);
                }
            });
        }

        function bindProj(json) {
            var $select = $('#in_proj');
            $select.find('option').remove();
            $.each(json, function (key, value) {
                //console.log(`<option value="${value.proj_id}">${value.proj_name}</option>`);
                $select.append(`<option value="${value.proj_id}">${value.proj_name}</option>`);

            });

            $('#in_proj').select2({
                placeholder: "เลือกโครงการ",
                allowClear: true
            });
        }

        function searchData() {
            var $lob = $("#in_lob");
            if ($lob.val() === null) { return alert('เลือกสายงาน'); }

            var $proj = $("#in_proj");
            var in_proj = $proj.val() == null ? null : $proj.val();
            //console.log(in_proj);

            var $month = $("#in_month")
            //alert($month.val());
            if (!$month.val()) { return alert('เลือกเดือนที่ต้องการออกรายงาน'); }
            //else { console.log($month.val()); }

            var $rdregion = $("#inr_region");
            var val_region = $rdregion.is(':checked') ? $rdregion.val() : null;

            var $rdlowrise = $("#inr_lowrise");
            var val_lowrise = $rdlowrise.is(':checked') ? $rdlowrise.val() : null;

            var $rdall = $("#inr_all");
            var val_all = $rdall.is(':checked') ? $rdall.val() : null;

            if (val_region === null && val_lowrise === null && val_all === null) {
                return alert('เลือกรูปแบบรายงาน');
            }

            var type_rpt;
            if (val_region !== null) {
                type_rpt = "region";
            }
            if (val_lowrise !== null) {
                type_rpt = "lowrise";
            }
            if (val_all !== null) {
                type_rpt = "all";
            }

            var data_;
            //if (($lob.val() !== null) && ($month !== null) &&
            //    (!val_region !== null || val_lowrise !== null || val_all !== null)) {
            var sm = $month.val();
            var arr = sm.split('-');
            var in_month = (arr[1] + arr[0]);

            var $proj_type = $("#in_proj_type");
            var in_proj_type = $proj_type.val();

            data_ = '{in_lob: "' + $lob.val() + '", in_projid: "' + in_proj + '", in_period: "' + in_month + '", in_projtype: "' + in_proj_type + '", type_rpt: "' + type_rpt + '"}';
            //}
            console.log(data_);


            //202206
            $.ajax({
                url: 'ReportKpiISO.aspx/GenerateFileExcel',
                type: 'POST',
                data: data_,
                contentType: 'application/json; charset=utf-8',
                dataType: 'json',
                success: function (response) {
                    if (response) {

                        //window.location = response;
                        //// OR
                        window.location = response.d;
                        //window.open(response.d, '_blank');
                         
                    }//end if
                },
                failure: function (response) {
                    alert(response.d);
                }
            });

            //if ($rdregion.is(':checked'))
            //    alert($rdregion.val() + ": true!!");
            //else alert('NO!');

        }

    </script>

</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <%--<form id="excelDownloadForm" action="ReportKpiISO.aspx/DownloadExcel" method="post" target="_self" >
        <input type="hidden" name="fileName" id="fileName" />
    </form>--%>

    <div class="card card-custom gutter-b example example-compact">
        <div class="card-header">
            <h3 class="card-title">รายงานวัตถุประสงค์คุณภาพ (ISO) รับมอบ - คงค้าง - แจ้งซ่อม</h3>

        </div>
        <!--begin::Form-->
        <div class="form">
            <div class="card-body">
                <%-- สายงาน --%>
                <div class="form-group row">
                    <label class="col-form-label text-right col-lg-3 col-sm-12">สายงาน</label>
                    <div class="col-lg-4 col-md-9 col-sm-12">
                        <select class="form-control select2" id="in_lob"></select>
                    </div>
                </div>
                <%-- ประเภทโครงการ --%>
                <div class="form-group row">
                    <label class="col-form-label text-right col-lg-3 col-sm-12">ประเภทโครงการ</label>
                    <div class="col-lg-4 col-md-9 col-sm-12">
                        <select class="form-control select2" id="in_proj_type"></select>
                    </div>
                </div>
                <%-- โครงการ --%>
                <div class="form-group row">
                    <label class="col-form-label text-right col-lg-3 col-sm-12">โครงการ</label>
                    <div class="col-lg-4 col-md-9 col-sm-12">
                        <select class="form-control select2" id="in_proj" multiple="multiple"></select>
                    </div>
                </div>
                <%-- ช่วงเวลา --%>
                <div class="form-group row">
                    <label class="col-form-label text-right col-lg-3 col-sm-12">ช่วงเวลา</label>
                    <div class="col-lg-4 col-md-9 col-sm-12">
                        <div class="input-group date" data-target-input="nearest">
                            <input type="text" class="form-control datetimepicker-input" id="in_month" autocomplete="off" placeholder="เลือกเดือนที่ต้องการออกรายงาน">
                            <div class="input-group-append" data-toggle="datetimepicker">
                                <span class="input-group-text">
                                    <i class="ki ki-calendar"></i>
                                </span>
                            </div>
                        </div>
                    </div>
                </div>
                <%-- รูปแบบรายงาน --%>
                <div class="form-group row">
                    <label class="col-form-label text-right col-lg-3 col-sm-12">รูปแบบรายงาน</label>
                    <div class="col-lg-4 col-md-9 col-sm-12">
                        <div class="radio-inline">
                            <label class="radio radio-success">
                                <input type="radio" name="radios5" id="inr_region" value="rpt_type_region"/>
                                <span></span>
                                รูปแบบภูมิภาค
                            </label>
                            <label class="radio radio-success">
                                <input type="radio" name="radios5" id="inr_lowrise" value="rpt_type_low_rise"/>
                                <span></span>
                                รูปแบบแนวราบ
                            </label>
                            <label class="radio radio-success">
                                <input type="radio" name="radios5" id="inr_all" value="rpt_type_all"/>
                                <span></span>
                                รูปแบบภูมิภาค (แนวราบ + อาคารสูง)
                            </label>
                        </div>
                    </div>
                </div>
                <%-- ออกรายงาน --%>
                <div class="form-group row">
                    <div class="col-lg-9 ml-lg-auto">
                        <input type="button" class="btn btn-success mr-2" value="ออกรายงาน" onclick="searchData()"/>
                        <%-- <input type="button" class="btn btn-secondary" value="ล้างค่า" /> --%>
                    </div>
                </div>
            </div>

            <%--<div class="card-footer">
                
            </div>--%>
        </div>
        <!--end::Form-->

    </div>
</asp:Content>